<#
.SYNOPSIS
    Create reports on computer status and health
.PARAMETER NoArchive
    Do not generate the zip archive. This is useful if you want to generate additional reports after this script.
.PARAMETER OnlyArchive
    Only create the archive from existing reports. This is useful if you have generated additional reports after this script.
.PARAMETER NoPerformanceDiagnostics
    Do not collect the live performance diagnostics. This makes the script considerably faster.
.PARAMETER PerformanceDuration
    How many seconds to sample the live performance diagnostics for.
    Run the report while the computer is slow to get useful data.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSReviewUnusedParameter", "Elevated", Justification="Used in utils")]
param(
    [switch]$Elevated,
    [switch]$NoArchive,
    [switch]$OnlyArchive,
    [switch]$NoPerformanceDiagnostics,
    [ValidateRange(5, 86400)][int]$PerformanceDuration = 180
)

Set-StrictMode -Version 3.0
. "${PSScriptRoot}\Utils.ps1"

if ($RepoInUserDir) {
    Update-Repo
}
Elevate($myinvocation.MyCommand.Definition)
if (! $RepoInUserDir) {
    Update-Repo
}

$host.ui.RawUI.WindowTitle = "Mika's reporting script"

# These variables are defined already here so that they are available also
# when only running the Compress-ReportArchive below.
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

Show-Output "Creating report of the display configuration."
Get-DisplayTopology | Format-List > "${Reports}\displays.txt"

Show-Output "Creating report of SSD/HDD SMART data."
Get-Disk | Get-StorageReliabilityCounter | Select-Object -Property "*" > "${Reports}\smart.txt"

Show-Output "Creating report of virtualization-based security."
Get-VirtualizationSecurityStatus | Format-List > "${Reports}\virtualization_based_security.txt"

Show-Output "Creating report of network configuration."
Get-NetIPConfiguration | Select-Object `
    "InterfaceAlias",
    "InterfaceDescription",
    @{n="MacAddress"; e={$_.NetAdapter.MacAddress}},
    @{n="MacAddressColon"; e={$_.NetAdapter.MacAddress.Replace("-", ":")}},
    @{n="IPv4Address"; e={$_.IPv4Address -join ", "}},
    @{n="IPv6Address"; e={$_.IPv6Address -join ", "}},
    @{n="IPv4DefaultGateway"; e={$_.IPv4DefaultGateway.NextHop -join ", "}},
    @{n="IPv6DefaultGateway"; e={$_.IPv6DefaultGateway.NextHop -join ", "}},
    @{n="DNSServer"; e={$_.DNSServer.ServerAddresses -join ", "}} > "${Reports}\network_configuration.txt"

Show-Output "Creating report of Plug and Play devices."
Get-PnPDevice > "${Reports}\pnp_devices.txt"

Show-Output "Extracting Windows Update logs."
Get-WindowsUpdateLog -LogPath "${Reports}\WindowsUpdate.log"

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

if (Test-CommandExists "dsregcmd") {
    Show-Output "Creating report of Microsoft Entra ID device registration."
    dsregcmd /status > "${Reports}\dsregcmd.txt"
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

$OpenVPNLogs = "${UserDir}\OpenVPN\log"
if (Test-Path "${OpenVPNLogs}") {
    Show-Output "OpenVPN log folder found. Copying logs."
    New-Item -Path "${Reports}" -Name "OpenVPN" -ItemType "directory" -Force | Out-Null
    Copy-Item -Path "${OpenVPNLogs}\*" -Destination "${Reports}\OpenVPN"
} else {
    Show-Output "OpenVPN log folder was not found."
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
# Live performance diagnostics
# -----
# This is last, because it takes several minutes and samples the running system.
# Everything above is a static snapshot, so the sampling is not disturbed by it.
if (-not $NoPerformanceDiagnostics) {
    Show-Output "Collecting live performance diagnostics for ${PerformanceDuration} seconds."
    Show-Output "If the computer is currently slow, leave it running and reproduce the slowness now."
    & "${PSScriptRoot}\Get-PerformanceDiagnostics.ps1" `
        -Duration $PerformanceDuration `
        -OutputPath "${Reports}\Performance"
} else {
    Show-Output "Skipping the live performance diagnostics."
}

# -----
# Packaging
# -----
if (-not $NoArchive) {
    Compress-ReportArchive
    Show-Output "The reporting script is ready." -ForegroundColor Green
    Show-Output "The reports can be found in the zip file on your desktop, and at `"${RepoPath}\Reports`"." -ForegroundColor Green
    Show-Output "If Mika requested you to run this script, please send the zip file from your desktop to him." -ForegroundColor Green
    Show-Output "You can close this window now." -ForegroundColor Green
}
