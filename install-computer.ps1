<#
.SYNOPSIS
    Install all the software and configuration usually needed for a computer.
#>

$GlobalWidth = 300;

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
    $List.Height = $Options.Count * 18;
    return $List;
}

# Import Windows Forms Assembly
Add-Type -AssemblyName System.Windows.Forms;
# Create a Form
$Form = New-Object -TypeName System.Windows.Forms.Form;
$Form.Text = "Mika's installer script"
$Form.Width = $GlobalWidth + 30;

$Table = New-Object System.Windows.Forms.TableLayoutPanel
$Table.RowCount = 2;
$Table.Width = $GlobalWidth + 10;
$Table.Height = 500
$Form.Controls.Add($Table);


$ChocolateyPrograms = [ordered]@{
    "Firefox" = "firefox"
    "LibreOffice" = "libreoffice";
    "OBS Studio" = "obs-studio";
    "OpenVPN" = "openvpn";
    "Steam" = "steam";
    "VLC" = "vlc";
}
$ChocolateyProgramsListBox = CreateList $Table "Centrally updated programs (Chocolatey)" $ChocolateyPrograms.keys;

# Show the form
$Form.ShowDialog();
# $Form.Close();

Write-Host "foo";
foreach($item in $ChocolateyProgramsListBox.SelectedItems){
    Write-Host $item.text;
    foreach($i in $item.SubItems){
        Write-Host $i.Text
    }
}
