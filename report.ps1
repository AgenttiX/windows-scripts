<#
.SYNOPSIS
    Create reports on computer status and health
#>

. "./utils.ps1"
Elevate($myinvocation.MyCommand.Definition)

$host.ui.RawUI.WindowTitle = "Mika's reporting script"

New-Item -Path "." -Name "downloads" -ItemType "directory" -Force
New-Item -Path "." -Name "reports" -ItemType "directory" -Force

Write-Host "Creating DirectX reports"
dxdiag /x ".\reports\dxdiag.xml"
dxdiag /t ".\reports\dxdiag.txt"
dxdiag /x ".\reports\dxdiag-whql.xml" /whql:on
dxdiag /t ".\reports\dxdiag-whql.txt" /whql:on

Write-Host "Creating battery report"
powercfg /batteryreport /output ".\reports\battery.html"

Write-Host "Creating WiFi report"
netsh wlan show wlanreport
cp "C:\ProgramData\Microsoft\Windows\WlanReport\wlan-report-latest.html" ./reports

Write-Host "Installing Geekbench"
$GeekbenchVersions = @("5.4.1", "4.4.4", "3.4.4", "2.4.3")
foreach ($Version in $GeekbenchVersions) {
    $Filename = "Geekbench-4.3.3-WindowsSetup.exe"
    $Url = "https://cdn.geekbench.com/$Filename"
}

Write-Host "Creating the report archive"
Compress-Archive -Path ".\reports" -DestinationPath ".\reports.zip" -CompressionLevel Optimal

Write-Host "The reporting script is ready. You can close this window now." -ForegroundColor Green
Write-Host 'The reports can be found in the "reports" subfolder, and in the corresponding zip file.' -ForegroundColor Green
Write-Host "If Mika requested you to run this script, please send the reports.zip file to him." -ForegroundColor Green
