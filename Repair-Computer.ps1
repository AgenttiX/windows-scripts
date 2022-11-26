<#
.SYNOPSIS
    Fix various issues with Windows
#>

. "./utils.ps1"
Elevate($MyInvocation.MyCommand.Definition)

Start-Transcript -Path "${LogPath}\Repair-Computer_$(Get-Date -Format "yyyy-MM-dd_HH-mm").txt"

Show-Output -ForegroundColor Cyan "Running Mika's repair script."

Show-Output -ForegroundColor Cyan "Running Windows System File Checker (SFC)."
sfc /scannow

Show-Output -ForegroundColor Cyan "Running DISM scan."
Dism /Online /Cleanup-Image /ScanHealth

Show-Output -ForegroundColor Cyan "Running DISM check."
Dism /Online /Cleanup-Image /CheckHealth

$Reply = Read-Host -Prompt "Do you want to run DISM repair? If no errors were found by the DISM scans above, this is not needed."
if ( $Reply -match "[yY]" ) {
    Show-Output -ForegroundColor Cyan "Running DISM repair."
    Dism /Online /Cleanup-Image /RestoreHealth
} else {
    Show-Output -ForegroundColor Cyan "Interpreting answer as a no. Skipping DISM repair."
}

Show-Output -ForegroundColor Cyan "Running CHKDSK."
chkdsk /R

Show-Output -ForegroundColor Cyan "Running Windows memory diagnostics."
Start-Process -NoNewWindow MdSched

Show-Output -ForegroundColor Green "The repair script is ready."
Stop-Transcript
