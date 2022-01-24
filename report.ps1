<#
.SYNOPSIS
    Create reports on computer status and health
.PARAMETER NoArchive
    Do not generate the zip archive. This is useful if you want to generate additional reports after this script.
.PARAMETER OnlyArchive
    Only create the archive from existing reports. This is useful if you have generated additional reports after this script.
#>

param(
    [switch]$NoArchive
)

. ".\utils.ps1"
Elevate($myinvocation.MyCommand.Definition)

$host.ui.RawUI.WindowTitle = "Mika's reporting script"

$Downloads = ".\downloads"
$Reports = ".\reports"

function Create-ReportArchive {
    Show-Output "Creating the report archive"
    $Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
    Compress-Archive -Path "${Reports}" -DestinationPath ".\reports_${Timestamp}.zip" -CompressionLevel Optimal
}

if ($OnlyArchive) {
    Create-ReportArchive
    exit
}


Show-Output "Running Mika's reporting script"
New-Item -Path "." -Name "downloads" -ItemType "directory" -Force
New-Item -Path "." -Name "reports" -ItemType "directory" -Force

Show-Output "Removing old reports"
Get-ChildItem "${Reports}/*" -Recurse | Remove-Item

Show-Output "Checking Windows Experience Index"
Get-CimInstance Win32_WinSat > "${Reports}\windows_experience_index.txt"

Show-Output "Creating DirectX reports"
dxdiag /x "${Reports}\dxdiag.xml"
dxdiag /t "${Reports}\dxdiag.txt"
dxdiag /x "${Reports}\dxdiag-whql.xml" /whql:on
dxdiag /t "${Reports}\dxdiag-whql.txt" /whql:on

Show-Output "Creating battery report"
powercfg /batteryreport /output "${Reports}\battery.html"

Show-Output "Creating WiFi report"
netsh wlan show wlanreport
cp "C:\ProgramData\Microsoft\Windows\WlanReport\wlan-report-latest.html" "${Reports}"

Show-Output "Creating report of installed Windows Store apps."
Get-AppxPackage > "${Reports}\appx_packages.txt"

$PTS = "${Env:SystemDrive}\phoronix-test-suite\phoronix-test-suite.bat"
if (Test-Path $PTS) {
    Show-Output "Creating Phoronix Test Suite (PTS) reports"
    & "$PTS" diagnostics > "${Reports}\pts_diagnostics.txt"
    & "$PTS" system-info > "${Reports}\pts_system_info.txt"
    & "$PTS" system-properties > "${Reports}\pts_system_properties.txt"
    & "$PTS" system-sensors > "${Reports}\pts_system_sensors.txt"
    & "$PTS" network-info > "${Reports}\pts_network_info.txt"
}

if (-not $NoArchive) {
    Create-ReportArchive
    Show-Output "The reporting script is ready." -ForegroundColor Green
    Show-Output 'The reports can be found in the "reports" subfolder, and in the corresponding zip file.' -ForegroundColor Green
    Show-Output "If Mika requested you to run this script, please send the reports.zip file to him." -ForegroundColor Green
    Show-Output "You can close this window now." -ForegroundColor Green
}
