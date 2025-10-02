<#
.SYNOPSIS
    Create reports on computer status and health
.PARAMETER NoArchive
    Do not generate the zip archive. This is useful if you want to generate additional reports after this script.
.PARAMETER OnlyArchive
    Only create the archive from existing reports. This is useful if you have generated additional reports after this script.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSReviewUnusedParameter", "Elevated", Justification="Used in utils")]
param(
    [switch]$Elevated,
    [switch]$NoArchive,
    [switch]$OnlyArchive
)

. "${PSScriptRoot}\Utils.ps1"
Elevate($myinvocation.MyCommand.Definition)

$host.ui.RawUI.WindowTitle = "Mika's reporting script"

# $Downloads = ".\Downloads"
$Reports = "${PSScriptRoot}\Reports"
$Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm"

function Compress-ReportArchive {
    Show-Output "Creating the report archive."
    Compress-Archive -Path "${Reports}" -DestinationPath "${DesktopPath}\IT_report_${Timestamp}.zip" -CompressionLevel Optimal
}

if ($OnlyArchive) {
    Compress-ReportArchive
    exit
}

# -----
# Initialization
# -----
Show-Output "Running Mika's reporting script."
New-Item -Path "." -Name "Reports" -ItemType "directory" -Force | Out-Null

Show-Output "Removing old reports."
Get-ChildItem "${Reports}/*" -Recurse | Remove-Item

Show-Output "Adding a README to the report."
(Get-Content "${PSScriptRoot}/Report-Readme-Template.txt").Replace("HOST", "${env:ComputerName}").Replace("TIMESTAMP", "${Timestamp}") | Set-Content "${Reports}\README.txt"

# -----
# Getter commands (in alphabetical order)
# -----
Show-Output "Creating report of installed Windows Store apps."
Get-AppxPackage > "${Reports}\appx_packages.txt"

Show-Output "Checking Windows Experience Index."
Get-CimInstance Win32_WinSat > "${Reports}\windows_experience_index.txt"

Show-Output "Creating report of basic computer info."
Get-ComputerInfo > "${Reports}\computer_info.txt"

Show-Output "Creating report of SSD/HDD SMART data."
Get-Disk | Get-StorageReliabilityCounter | Select-Object -Property "*" > "${Reports}\smart.txt"

Show-Output "Creating report of Plug and Play devices."
Get-PnPDevice > "${Reports}\pnp_devices.txt"

# -----
# External commands (in alphabetical order)
# -----
if (Test-CommandExists "choco") {
    Show-Output "Creating report of installed Chocolatey apps."
    choco --local > "${Reports}\choco.txt"
} else {
    Show-Output "The command `"choco`" was not found."
}

if (Test-CommandExists "dxdiag") {
    Show-Output "Creating DirectX reports."
    dxdiag /x "${Reports}\dxdiag.xml"
    dxdiag /t "${Reports}\dxdiag.txt"
    dxdiag /x "${Reports}\dxdiag-whql.xml" /whql:on
    dxdiag /t "${Reports}\dxdiag-whql.txt" /whql:on
} else {
    Show-Output "The command `"dxdiag`" was not found."
}

if (Test-CommandExists "gpresult") {
    Show-Output "Creating report of group policies."
    gpresult /h "${Reports}\gpresult.html" /f
} else {
    Show-Output "The command `"gpresult`" was not found."
}

if (Test-CommandExists "manage-bde") {
    manage-bde -status > "${Reports}\manage-bde.txt"
} else {
    Show-Output "The command `"manage-bde`" was not found."
}

if (Test-CommandExists "netsh") {
    Show-Output "Creating WLAN report."
    netsh wlan show wlanreport
    $WlanReportPath1 = "C:\ProgramData\Microsoft\Windows\WlanReport\wlan-report-latest.html"
    $WlanReportPath2 = "C:\ProgramData\Microsoft\Windows\WlanReport\wlan_report_latest.html"
    if (Test-Path "${WlanReportPath1}") {
        Copy-Item "${WlanReportPath1}" "${Reports}"
    } elseif (Test-path "${WlanReportPath2}") {
        Copy-Item "${WlanReportPath2}" "${Reports}"
    } else {
        Show-Output -ForegroundColor Red "The WLAN report was not found."
    }
} else {
    Show-Output "The command `"netsh`" was not found."
}

if (Test-CommandExists "powercfg") {
    Show-Output "Creating battery report."
    powercfg /availablesleepstates > "${Reports}\powercfg_sleepstates.html"
    powercfg /batteryreport /output "${Reports}\powercfg_battery.html"
    powercfg /devicequery wake_armed > "${Reports}\powercfg_devicequery_wake_armed.txt"
    powercfg /energy /output "${Reports}\powercfg_energy.html"
    powercfg /getactivescheme > "${Reports}\powercfg_activescheme.txt"
    powercfg /lastwake > "${Reports}\powercfg_lastwake.txt"
    powercfg /list > "${Reports}\powercfg_list.txt"
    powercfg /provisioningxml /output "${Reports}\powercfg_provisioning.xml"
    powercfg /sleepstudy /output "${Reports}\powercfg_sleepstudy.html"
    powercfg /srumutil /output "${Reports}\powercfg_srumutil.csv" /csv
    powercfg /systempowerreport /output "${Reports}\powercfg_systempowerreport.html"
    powercfg /waketimers > "${Reports}\powercfg_waketimers.txt"
} else {
    Show-Output "The command `"powercfg`" was not found."
}

# -----
# Complex external programs
# -----
$PTS = "${Env:SystemDrive}\phoronix-test-suite\phoronix-test-suite.bat"
if (Test-Path $PTS) {
    Show-Output "Creating Phoronix Test Suite (PTS) reports"
    & "$PTS" diagnostics > "${Reports}\pts_diagnostics.txt"
    & "$PTS" system-info > "${Reports}\pts_system_info.txt"
    & "$PTS" system-properties > "${Reports}\pts_system_properties.txt"
    & "$PTS" system-sensors > "${Reports}\pts_system_sensors.txt"
    & "$PTS" network-info > "${Reports}\pts_network_info.txt"
} else {
    Show-Output "Phoronix Test Suite (PTS) was not found."
}

# -----
# Packaging
# -----
if (-not $NoArchive) {
    Compress-ReportArchive
    Show-Output "The reporting script is ready." -ForegroundColor Green
    Show-Output "The reports can be found in the zip file on your desktop, and at `"${RepoPath}\Reports`"." -ForegroundColor Green
    Show-Output "If Mika requested you to run this script, please send the zip file to him." -ForegroundColor Green
    Show-Output "You can close this window now." -ForegroundColor Green
}
