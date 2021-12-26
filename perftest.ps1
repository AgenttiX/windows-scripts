<#
.SYNOPSIS
    Test computer performance, stability and capabilities.
    Warning! The Phoronix Test Suite tests used by this script will require over 20 GB of disk space!
.PARAMETER All
    Download all tests, even the manual ones.
.PARAMETER DownloadOnly
    Only download the tests, but don't run them.
.PARAMETER Geekbench
    Download Geekbench (manual due to licensing)
.PARAMETER Unigine
    Download Unigine benchmarks (manual due to licensing)
#>

param(
    [switch]$All,
    [switch]$DownloadOnly,
    [switch]$Geekbench,
    [switch]$Unigine
)

. "./utils.ps1"
Elevate($MyInvocation.MyCommand.Definition)

Write-Host "Running Mika's performance testing script"
$ScriptPath = $MyInvocation.MyCommand.Path
$RepoPath = Split-Path $ScriptPath -Parent
$StartPath = Get-Location
# The downloads folder has to be created already here for PTS download
New-Item -Path "$RepoPath" -Name "downloads" -ItemType "directory" -Force
$Downloads = "${RepoPath}\downloads"

$PTS = "${Env:SystemDrive}\phoronix-test-suite\phoronix-test-suite.bat"
if (-Not (Test-Path "$PTS")) {
    Write-Host "Downloading Phoronix Test Suite (PTS)"
    $PTS_version = "10.8.0"
    Invoke-WebRequest -Uri "https://github.com/phoronix-test-suite/phoronix-test-suite/archive/v${PTS_version}.zip" -OutFile "${Downloads}\phoronix-test-suite-${PTS_version}.zip"
    Write-Host "Extracting Phoronix Test Suite (PTS)"
    Expand-Archive -LiteralPath "${Downloads}\phoronix-test-suite-${PTS_version}.zip" -DestinationPath "${Downloads}\phoronix-test-suite-${PTS_version}" -Force
    # The installation script needs to be executed in its directory
    cd "${Downloads}\phoronix-test-suite-${PTS_version}\phoronix-test-suite-${PTS_version}"
    Write-Host "Installing Phoronix Test Suite (PTS)."
    Write-Host "You may get prompts asking whether you accept the EULA and whether to send anonymous usage statistics."
    Write-Host "Please select yes to the former and preferably to the latter as well."
    & ".\install.bat"
    cd "$StartPath"
    if (-Not (Test-Path "$PTS")) {
        Write-Host "Phoronix Test Suite (PTS) installation failed."
        exit 1
    }
    Write-Host "Phoronix Test Suite (PTS) has been installed. It is highly recommended that you log in now so that you can manage the uploaded results."
    & "$PTS" openbenchmarking-login
    Write-Host "You must now configure the batch mode according to your preferences. Some of the settings may get overwritten later by this script."
    & "$PTS" batch-setup
}
& "$PTS" openbenchmarking-refresh

# The reporting has to be after PTS installation to be able to generate the PTS reports
& ".\report.ps1"
# This folder is created by the reporting script
$Reports = ".\reports"


Write-Host "Installing dependencies"
Install-Chocolatey

Write-Host "Installing tests"
choco upgrade -y speedtest
if ($All -or $Unigine) {
    choco upgrade -y heaven-benchmark superposition-benchmark valley-benchmark
}

if (-not (Test-Path "${Downloads}\UserBenchMark.exe")) {
    Write-Host "Downloading UserBenchMark"
    Invoke-WebRequest -Uri "https://www.userbenchmark.com/resources/download/UserBenchMark.exe" -OutFile "${Downloads}\UserBenchMark.exe"
}

# Invoke-WebRequest -Uri "https://github.com/microsoft/diskspd/releases/download/v2.1/DiskSpd.ZIP" -OutFile "${Downloads}\DiskSpd.ZIP"
# Expand-Archive -LiteralPath "${Downloads}\DiskSpd.ZIP" -DestinationPath "${Downloads}\DiskSpd"
# https://github.com/ayavilevich/DiskSpdAuto

if ($Geekbench -or $All) {
    $GeekbenchVersions = @("5.4.4", "4.4.4", "3.4.4", "2.4.3")
    foreach ($Version in $GeekbenchVersions) {
        Write-Host "Downloading Geekbench ${Version}"
        $Filename = "Geekbench-$Version-WindowsSetup.exe"
        $Url = "https://cdn.geekbench.com/$Filename"
        Invoke-WebRequest -Uri "$Url" -OutFile "$Downloads\$Filename"
    }
    foreach ($Version in $GeekbenchVersions) {
        Write-Host "Installing Geekbench ${Version}"
        Start-Process -NoNewWindow -Wait "${Downloads}\Geekbench-${Version}-WindowsSetup.exe"
    }
}

if ($DownloadOnly) {
    Write-Host "-DownloadOnly was specified. Everything is downloaded, so exiting now."
    exit
}

Write-Host "Configuring PTS"
$Env:MONITOR="all"
$Env:PERFORMANCE_PER_WATT="1"
& "$PTS" user-config-set AllowResultUploadsToOpenBenchmarking=TRUE
& "$PTS" user-config-set AlwaysUploadSystemLogs=TRUE
& "$PTS" user-config-set AlwaysUploadResultsToOpenBenchmarking=TRUE
& "$PTS" user-config-set AnonymousUsageReporting=TRUE
& "$PTS" user-config-set SaveSystemLogs=TRUE
& "$PTS" user-config-set SaveResults=TRUE
# & "$PTS" user-config-set PromptForTestDescription=FALSE
& "$PTS" user-config-set PromptSaveName=FALSE

Write-Host "Running PTS"
& "$PTS" batch-benchmark `
    pts/av1 `
    pts/compiler `
    pts/compression `
    pts/cryptocurrency `
    pts/cryptography `
    pts/disk `
    pts/hpc `
    pts/machine-learning `
    pts/memory `
    pts/opencl `
    pts/python `
    pts/scientific-computing `
    pts/video-encoding `
    pts/cinebench `
    pts/compress-7zip `
    pts/encode-flac `
    pts/intel-mlc `
    pts/x264

Write-Host "Running Speedtest"
speedtest --accept-license --accept-gdpr --format=csv --output-header > "${Reports}\speedtest.csv"

Write-Host "Running Windows performance monitoring"
Start-Process -NoNewWindow -Wait perfmon /report

Write-Host "Running winsat disk test"
winsat disk -drive C > "${Reports}\winsat_disk.txt"

# furmark /log_temperature /log_score

# Windows memory diagnostics
Write-Host "Running Windows memory diagnostics"
Start-Process -NoNewWindow -Wait MdSched

# Running manual tests from a script is rather pointless.
#
# Waiting for a GUI program to finish
# https://stackoverflow.com/a/7908022/
#
# Start-Process -NoNewWindow -Wait ".\downloads\UserBenchMark.exe"
#
# Start-Process -NoNewWindow -Wait shell:appsFolder\MAXONComputerGmbH.Cinebench_rsne5bsk8s7tj!App
#
# # 3DMark and PCMark 10 have command-line support only in their professional versions.
# # https://support.benchmarks.ul.com/support/solutions/articles/44002145411-run-3dmark-benchmarks-from-the-command-line
# # https://support.benchmarks.ul.com/support/solutions/articles/44002182309-run-pcmark-10-benchmarks-from-the-command-line
# $3DMarkPath = "${Env:ProgramFiles(x86)}\Steam\steamapps\common\3DMark\bin\x64\3DMark.exe"
# if (Test-Path "$3DMarkPath") {
#     Start-Process -NoNewWindow -Wait "$3DMarkPath"
# }
# $PCMarkPath = "${Env:ProgramFiles(x86)}\Steam\steamapps\common\PCMark 10\bin\x64\PCMark10.exe"
# if (Test-Path "$PCMarkPath") {
#     Start-Process -NoNewWindow -Wait $PCMarkPath
# }
