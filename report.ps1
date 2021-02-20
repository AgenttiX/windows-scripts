<#
.SYNOPSIS
    Create reports on computer status and health
#>

. "./utils.ps1"
Elevate($myinvocation.MyCommand.Definition)

New-Item -Path "." -Name "reports" -ItemType "directory" -Force

Write-Host "Creating DirectX reports"
dxdiag /x ".\reports\dxdiag.xml"
dxdiag /t ".\reports\dxdiag.txt"
dxdiag /x ".\reports\dxdiag-whql.xml" /whql:on
dxdiag /t ".\reports\dxdiag-whql.txt" /whql:on

Write-Host "Creating battery report"
powercfg /batteryreport /output "./reports/battery.html"

Write-Host "Creating WiFi report"
netsh wlan show wlanreport
cp "C:\ProgramData\Microsoft\Windows\WlanReport\wlan-report-latest.html" ./reports
