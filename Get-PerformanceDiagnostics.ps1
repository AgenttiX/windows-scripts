<#
.SYNOPSIS
    Collect live performance diagnostics to investigate system slowness.
.DESCRIPTION
    Samples performance counters over a period of time and combines the results with the
    system configuration and the Windows event log to find the causes of system slowness.

    This script is designed to be run while the slowness is happening, e.g. during a
    video conference. It is especially aimed at laptops, where the performance is often
    limited by the power and thermal budget shared by the CPU, the integrated GPU and
    the external displays.

    The results are written to a directory of text and CSV files, and a human-readable
    summary with the most important findings is both printed and saved.
.PARAMETER Duration
    The duration of the sampling in seconds.
.PARAMETER Interval
    The interval between samples in seconds.
.PARAMETER OutputPath
    The directory to write the results to.
    Defaults to a "Performance" subdirectory of the reports directory.
.PARAMETER EventHistoryDays
    How many days of event log history to summarize for context.
.PARAMETER NoHWiNFO
    Do not run HWiNFO, even if it is installed.
.PARAMETER HWiNFOPath
    The path of the HWiNFO executable. Autodetected if not given.
.EXAMPLE
    .\Get-PerformanceDiagnostics.ps1
    Sample for the default two minutes.
.EXAMPLE
    .\Get-PerformanceDiagnostics.ps1 -Duration 600 -Interval 5
    Sample for ten minutes, e.g. for the duration of a video conference.
.NOTES
    The ACPI thermal zone temperatures and the HWiNFO sensors are available only to
    administrators. This script does not elevate itself, because it is meant to be started
    at the moment the slowness occurs, and a UAC prompt would then disturb the user and
    distort the measurement. Run it from an administrator shell to get those as well.
#>

param(
    [ValidateRange(5, 86400)][int]$Duration = 120,
    [ValidateRange(1, 300)][int]$Interval = 2,
    [string]$OutputPath,
    [ValidateRange(1, 365)][int]$EventHistoryDays = 14,
    [switch]$NoHWiNFO,
    [string]$HWiNFOPath
)

Set-StrictMode -Version 3.0
. "${PSScriptRoot}\Utils.ps1"

if (-not $OutputPath) {
    $OutputPath = "${PSScriptRoot}\Reports\Performance"
}

# -----
# Helper functions
# -----

function Get-SampleStatistics {
    <#
    .SYNOPSIS
        Compute descriptive statistics for a set of samples.
    .PARAMETER Values
        The samples to compute the statistics for.
    #>
    [OutputType([PSCustomObject])]
    param(
        [AllowNull()][AllowEmptyCollection()][double[]]$Values
    )
    $Clean = @($Values | Where-Object { $null -ne $_ -and -not [double]::IsNaN($_) })
    if ($Clean.Count -eq 0) {
        return [PSCustomObject]@{ Count = 0; Min = $null; Median = $null; Mean = $null; P95 = $null; Max = $null }
    }
    $Sorted = @($Clean | Sort-Object)
    $Median = $Sorted[[math]::Floor(($Sorted.Count - 1) / 2)]
    if ($Sorted.Count % 2 -eq 0) {
        $Median = ($Sorted[$Sorted.Count / 2 - 1] + $Sorted[$Sorted.Count / 2]) / 2
    }
    # Nearest-rank 95th percentile
    $P95Index = [math]::Min($Sorted.Count - 1, [math]::Ceiling(0.95 * $Sorted.Count) - 1)
    if ($P95Index -lt 0) { $P95Index = 0 }
    return [PSCustomObject]@{
        Count = $Sorted.Count
        Min = [math]::Round($Sorted[0], 2)
        Median = [math]::Round($Median, 2)
        Mean = [math]::Round(($Sorted | Measure-Object -Average).Average, 2)
        P95 = [math]::Round($Sorted[$P95Index], 2)
        Max = [math]::Round($Sorted[$Sorted.Count - 1], 2)
    }
}

function ConvertTo-CounterKey {
    <#
    .SYNOPSIS
        Normalize a performance counter path so that it can be used as a lookup key.
    .DESCRIPTION
        Get-Counter returns paths that are prefixed with the computer name and are lowercased,
        so the requested paths cannot be matched to the results directly.
    .PARAMETER Path
        The counter path to normalize.
    #>
    [OutputType([string])]
    param(
        [Parameter(Mandatory=$true)][string]$Path
    )
    return ($Path -replace "^\\\\[^\\]+", "").ToLowerInvariant()
}

function Select-AvailableCounter {
    <#
    .SYNOPSIS
        Filter a set of performance counters down to those available on this system.
    .DESCRIPTION
        Performance counter names are localized on some Windows installations, and some
        counters are missing depending on the hardware. Unavailable counters would make
        the whole Get-Counter call fail, so they have to be filtered out in advance.
    .PARAMETER Counters
        An ordered dictionary of column name to counter path.
    #>
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
    param(
        [Parameter(Mandatory=$true)][System.Collections.Specialized.OrderedDictionary]$Counters
    )
    $Available = [ordered]@{}
    $Missing = @()
    foreach ($Name in $Counters.Keys) {
        try {
            $null = Get-Counter -Counter $Counters[$Name] -MaxSamples 1 -ErrorAction Stop
            $Available[$Name] = $Counters[$Name]
        } catch {
            $Missing += "${Name} ($($Counters[$Name]))"
        }
    }
    if ($Missing.Count -gt 0) {
        Show-Information -ForegroundColor Yellow "These performance counters are not available and will be skipped:"
        foreach ($Item in $Missing) { Show-Information "  - ${Item}" }
    }
    return $Available
}

function Get-GpuEngineUsage {
    <#
    .SYNOPSIS
        Get the current GPU utilization, grouped by engine type.
    .DESCRIPTION
        The video decode and encode engines are of particular interest, since video
        conferencing loads them heavily.
    #>
    [OutputType([hashtable])]
    param()
    $Result = @{ Total = 0.0; "3D" = 0.0; VideoDecode = 0.0; VideoEncode = 0.0; Copy = 0.0; Other = 0.0 }
    try {
        $Samples = (Get-Counter "\GPU Engine(*)\Utilization Percentage" -ErrorAction Stop).CounterSamples
    } catch {
        return $Result
    }
    foreach ($Sample in $Samples) {
        $Value = [double]$Sample.CookedValue
        if ($Value -le 0) { continue }
        $Result.Total += $Value
        if ($Sample.InstanceName -match "engtype_(\w+)") {
            switch -Regex ($Matches[1]) {
                "^3D$" { $Result["3D"] += $Value; break }
                "^VideoDecode" { $Result.VideoDecode += $Value; break }
                "^VideoEncode" { $Result.VideoEncode += $Value; break }
                "^Copy" { $Result.Copy += $Value; break }
                default { $Result.Other += $Value }
            }
        } else {
            $Result.Other += $Value
        }
    }
    foreach ($Key in @($Result.Keys)) { $Result[$Key] = [math]::Round($Result[$Key], 2) }
    return $Result
}

function Get-ThermalZoneTemperature {
    <#
    .SYNOPSIS
        Get the ACPI thermal zone temperatures in degrees Celsius.
    .NOTES
        This requires administrator privileges, and is not implemented by all firmware.
    #>
    [OutputType([System.Array])]
    param()
    try {
        return @(
            Get-CimInstance -Namespace "root/wmi" -ClassName "MSAcpi_ThermalZoneTemperature" -ErrorAction Stop |
                ForEach-Object {
                    [PSCustomObject]@{
                        Name = $_.InstanceName
                        Celsius = [math]::Round($_.CurrentTemperature / 10 - 273.15, 1)
                    }
                }
        )
    } catch {
        return @()
    }
}

function Get-ProcessorPowerEventSummary {
    <#
    .SYNOPSIS
        Summarize the processor power management events from the Windows event log.
    .DESCRIPTION
        Event 37 of Microsoft-Windows-Kernel-Processor-Power means that the system firmware
        is limiting the CPU speed. This is the signature of thermal or power limit throttling
        enforced by the embedded controller, as opposed to throttling by Windows.
    .PARAMETER StartTime
        The beginning of the time range to summarize.
    .PARAMETER EndTime
        The end of the time range to summarize.
    #>
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory=$true)][DateTime]$StartTime,
        [DateTime]$EndTime = (Get-Date)
    )
    $Events = @()
    try {
        $Events = @(Get-WinEvent -FilterHashtable @{
            LogName = "System"
            ProviderName = "Microsoft-Windows-Kernel-Processor-Power"
            Id = 37
            StartTime = $StartTime
            EndTime = $EndTime
        } -ErrorAction Stop)
    } catch {
        # Get-WinEvent throws if there are no matching events.
        $Events = @()
    }

    # Each event reports how many seconds one processor spent in a reduced performance state
    # since the previous report for that same processor. The seconds must therefore be summed
    # per processor. Summing them over all processors would multiply the real duration by the
    # number of affected processors.
    #
    # The event properties are used instead of the message text, since the message is
    # localized. Property 0 is the processor group, 1 is the processor number and
    # 2 is the number of seconds.
    $SecondsPerProcessor = @{}
    $LongestSingleReport = 0
    foreach ($Event in $Events) {
        $Group = $null
        $Processor = $null
        $Seconds = $null
        $Properties = @($Event.Properties)
        if ($Properties.Count -ge 3) {
            $Group = [int]$Properties[0].Value
            $Processor = [int]$Properties[1].Value
            $Seconds = [int]$Properties[2].Value
        } else {
            # Fall back to the message text if the event schema is not what we expect.
            $Message = $Event.Message -replace "\s+", " "
            if ($Message -match "processor (\d+) in group (\d+)") {
                $Processor = [int]$Matches[1]
                $Group = [int]$Matches[2]
            }
            if ($Message -match "for (\d+) seconds") { $Seconds = [int]$Matches[1] }
        }
        if ($null -eq $Processor -or $null -eq $Seconds) { continue }
        $Key = "${Group}:${Processor}"
        if (-not $SecondsPerProcessor.ContainsKey($Key)) { $SecondsPerProcessor[$Key] = 0 }
        $SecondsPerProcessor[$Key] += $Seconds
        if ($Seconds -gt $LongestSingleReport) { $LongestSingleReport = $Seconds }
    }

    $WorstSeconds = 0
    foreach ($Key in $SecondsPerProcessor.Keys) {
        if ($SecondsPerProcessor[$Key] -gt $WorstSeconds) { $WorstSeconds = $SecondsPerProcessor[$Key] }
    }

    return [PSCustomObject]@{
        EventCount = @($Events).Count
        # The throttled time of the worst affected processor
        WorstProcessorSeconds = $WorstSeconds
        # The longest continuous period reported by a single event
        LongestSingleReportSeconds = $LongestSingleReport
        AffectedProcessorCount = $SecondsPerProcessor.Keys.Count
        AffectedProcessors = (@($SecondsPerProcessor.Keys) | Sort-Object) -join ", "
        FirstEvent = $(if (@($Events).Count -gt 0) { $Events[@($Events).Count - 1].TimeCreated } else { $null })
        LastEvent = $(if (@($Events).Count -gt 0) { $Events[0].TimeCreated } else { $null })
    }
}

function Get-EventLogSummary {
    <#
    .SYNOPSIS
        Summarize the warnings and errors in an event log over a time range.
    .PARAMETER LogName
        The name of the event log.
    .PARAMETER StartTime
        The beginning of the time range.
    .PARAMETER EndTime
        The end of the time range.
    .PARAMETER MaxGroups
        The maximum number of event groups to return.
    #>
    [OutputType([System.Array])]
    param(
        [string]$LogName = "System",
        [Parameter(Mandatory=$true)][DateTime]$StartTime,
        [DateTime]$EndTime = (Get-Date),
        [int]$MaxGroups = 25
    )
    try {
        $Events = @(Get-WinEvent -FilterHashtable @{
            LogName = $LogName
            StartTime = $StartTime
            EndTime = $EndTime
            Level = 1, 2, 3
        } -ErrorAction Stop)
    } catch {
        return @()
    }
    return @(
        $Events | Group-Object -Property ProviderName, Id | Sort-Object Count -Descending |
            Select-Object -First $MaxGroups | ForEach-Object {
                $Message = ($_.Group[0].Message -replace "\s+", " ")
                if ($Message.Length -gt 200) { $Message = $Message.Substring(0, 200) + "..." }
                [PSCustomObject]@{
                    Count = $_.Count
                    Provider = $_.Group[0].ProviderName
                    Id = $_.Group[0].Id
                    Level = $_.Group[0].LevelDisplayName
                    Message = $Message
                }
            }
    )
}

function Get-TopProcessCpuUsage {
    <#
    .SYNOPSIS
        Compute how much CPU time each process consumed between two snapshots.
    .PARAMETER Before
        The snapshot taken before the measurement.
    .PARAMETER After
        The snapshot taken after the measurement.
    .PARAMETER ElapsedSeconds
        The wall clock time between the snapshots.
    .PARAMETER Count
        The number of processes to return.
    #>
    [OutputType([System.Array])]
    param(
        [Parameter(Mandatory=$true)][hashtable]$Before,
        [Parameter(Mandatory=$true)][hashtable]$After,
        [Parameter(Mandatory=$true)][double]$ElapsedSeconds,
        [int]$Count = 20
    )
    $LogicalProcessors = [Environment]::ProcessorCount
    $Results = New-Object System.Collections.Generic.List[PSObject]
    foreach ($Key in $After.Keys) {
        $EndEntry = $After[$Key]
        $StartCpu = 0.0
        if ($Before.ContainsKey($Key)) { $StartCpu = $Before[$Key].Cpu }
        $CpuDelta = $EndEntry.Cpu - $StartCpu
        if ($CpuDelta -le 0) { continue }
        $Results.Add([PSCustomObject]@{
            Name = $EndEntry.Name
            Id = $EndEntry.Id
            CpuSeconds = [math]::Round($CpuDelta, 1)
            # The percentage of one logical processor
            CpuPercent = [math]::Round(100 * $CpuDelta / $ElapsedSeconds, 1)
            # The percentage of the whole CPU
            SystemCpuPercent = [math]::Round(100 * $CpuDelta / ($ElapsedSeconds * $LogicalProcessors), 2)
            WorkingSetMB = $EndEntry.WorkingSetMB
        })
    }
    return @($Results | Sort-Object CpuSeconds -Descending | Select-Object -First $Count)
}

function Get-ProcessSnapshot {
    <#
    .SYNOPSIS
        Take a snapshot of the CPU time and memory usage of all processes.
    #>
    [OutputType([hashtable])]
    param()
    $Snapshot = @{}
    foreach ($Process in (Get-Process -ErrorAction SilentlyContinue)) {
        try {
            $Cpu = 0.0
            if ($null -ne $Process.CPU) { $Cpu = [double]$Process.CPU }
            # The process ID is reused, so the start time is needed to identify a process.
            $Key = "$($Process.Id)"
            $Snapshot[$Key] = [PSCustomObject]@{
                Name = $Process.Name
                Id = $Process.Id
                Cpu = $Cpu
                WorkingSetMB = [math]::Round($Process.WorkingSet64 / 1MB, 0)
            }
        } catch {
            # Processes that exit while being enumerated cannot be read.
            continue
        }
    }
    return $Snapshot
}

function Find-HWiNFO {
    <#
    .SYNOPSIS
        Find the HWiNFO executable.
    #>
    [OutputType([string])]
    param(
        [string]$Path
    )
    if ($Path) {
        if (Test-Path -LiteralPath $Path) { return $Path }
        Show-Information -ForegroundColor Yellow "HWiNFO was not found at the given path `"${Path}`"."
        return ""
    }
    $Candidates = @(
        "${env:ProgramFiles}\HWiNFO64\HWiNFO64.EXE",
        "${env:ProgramFiles}\HWiNFO32\HWiNFO32.EXE",
        "${env:ProgramFiles(x86)}\HWiNFO64\HWiNFO64.EXE",
        "${env:ProgramFiles(x86)}\HWiNFO32\HWiNFO32.EXE"
    )
    foreach ($Candidate in $Candidates) {
        if ($Candidate -and (Test-Path -LiteralPath $Candidate)) { return $Candidate }
    }
    return ""
}

function Start-HWiNFOLogging {
    <#
    .SYNOPSIS
        Start logging the HWiNFO sensors to a CSV file.
    .DESCRIPTION
        HWiNFO reads the sensors that Windows does not expose at all, most importantly the
        CPU package power and the firmware performance limit reasons. These tell apart
        thermal throttling from power limit throttling.

        Command line logging is a feature of the paid HWiNFO Pro version, and HWiNFO
        requires administrator privileges. This function therefore verifies that the log
        file actually appears, and gives up gracefully if it does not.
    .PARAMETER ExecutablePath
        The path of the HWiNFO executable.
    .PARAMETER LogPath
        The path of the CSV file to write.
    .PARAMETER PollRateMs
        The sensor polling period in milliseconds.
    .OUTPUTS
        The HWiNFO process object, or $null if the logging could not be started.
    .LINK
        https://www.hwinfo.com/forum/threads/supported-command-line-parameters-in-hwinfo64-pro.8549/
    #>
    [OutputType([System.Diagnostics.Process])]
    param(
        [Parameter(Mandatory=$true)][string]$ExecutablePath,
        [Parameter(Mandatory=$true)][string]$LogPath,
        [int]$PollRateMs = 2000
    )

    if (-not (Test-Admin)) {
        Show-Information -ForegroundColor Yellow "HWiNFO requires administrator privileges. Skipping the sensor logging."
        return $null
    }
    if (Get-Process -Name "HWiNFO64", "HWiNFO32" -ErrorAction SilentlyContinue) {
        Show-Information -ForegroundColor Yellow (
            "HWiNFO is already running. Only one instance can run at a time, " +
            "so the sensor logging is skipped. Close HWiNFO and run this script again."
        )
        return $null
    }
    if (Test-Path -LiteralPath $LogPath) { Remove-Item -LiteralPath $LogPath -Force }

    Show-Information "Starting the HWiNFO sensor logging."
    $Process = $null
    try {
        # There must be no space between -l and the file name.
        # -log_format=1 produces a single header row, which is far easier to parse.
        $Process = Start-Process `
            -FilePath $ExecutablePath `
            -ArgumentList "-l`"${LogPath}`"", "-poll_rate=${PollRateMs}", "-log_format=1" `
            -PassThru `
            -ErrorAction Stop
    } catch {
        Show-Information -ForegroundColor Yellow "Could not start HWiNFO: $($_.Exception.Message)"
        return $null
    }

    # Wait for the log file to appear, so that we can tell whether the logging works at all.
    $Deadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $Deadline) {
        if (Test-Path -LiteralPath $LogPath) {
            Show-Information "HWiNFO is logging the sensors to `"${LogPath}`"."
            return $Process
        }
        Start-Sleep -Milliseconds 500
    }

    Show-Information -ForegroundColor Yellow (
        "HWiNFO did not create a log file. Command line sensor logging requires the paid " +
        "HWiNFO Pro version, so this is expected with the free version. Continuing without it."
    )
    Stop-HWiNFOLogging -Process $Process
    return $null
}

function Stop-HWiNFOLogging {
    <#
    .SYNOPSIS
        Stop HWiNFO and let it flush its log file.
    .DESCRIPTION
        HWiNFO has to be closed gracefully, so that it writes the trailing rows of the CSV.
        Stop-Process terminates the process outright and may leave the file unflushed,
        so taskkill without /F is used to request a clean shutdown instead.
    .PARAMETER Process
        The HWiNFO process to stop.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        "PSUseShouldProcessForStateChangingFunctions",
        "",
        Justification="Stops a process that this script started itself"
    )]
    [OutputType([void])]
    param(
        [Parameter(Mandatory=$true)][AllowNull()][System.Diagnostics.Process]$Process
    )
    if ($null -eq $Process) { return }
    Show-Output "Stopping the HWiNFO sensor logging."
    if ($Process.HasExited) { return }
    try {
        $null = & taskkill.exe /PID $Process.Id 2>&1
    } catch {
        Show-Output -ForegroundColor Yellow "Could not stop HWiNFO gracefully: $($_.Exception.Message)"
    }
    if (-not $Process.WaitForExit(10000)) {
        Show-Output -ForegroundColor Yellow "HWiNFO did not exit gracefully. Terminating it."
        Stop-Process -Id $Process.Id -Force -ErrorAction SilentlyContinue
    }
}

# -----
# Initialization
# -----

# This script deliberately does not elevate itself. It is meant to be started at the moment
# the slowness occurs, and a UAC prompt at that moment would both disturb the user and
# distort the measurement. Everything except the thermal zone temperatures and the HWiNFO
# sensors works without elevation. New-Report.ps1 has already elevated when it calls this.

$host.ui.RawUI.WindowTitle = "Mika's performance diagnostics script"
Show-Output "Running Mika's performance diagnostics script."

New-Item -Path $OutputPath -ItemType "Directory" -Force | Out-Null
$Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
$SummaryPath = "${OutputPath}\performance_summary.txt"
$SamplesPath = "${OutputPath}\performance_samples.csv"
$ProcessesPath = "${OutputPath}\performance_processes.csv"
$EventsPath = "${OutputPath}\performance_events.txt"

if (-not (Test-Admin)) {
    Show-Output -ForegroundColor Yellow (
        "Not running as an administrator. The thermal zone temperatures and some " +
        "HWiNFO sensors will not be available."
    )
}

# -----
# Static configuration snapshot
# -----

Show-Output "Collecting the system configuration."

$ComputerSystem = Get-CimInstance -ClassName "Win32_ComputerSystem"
$OperatingSystem = Get-CimInstance -ClassName "Win32_OperatingSystem"
$Processor = @(Get-CimInstance -ClassName "Win32_Processor")[0]
$VideoControllers = @(Get-CimInstance -ClassName "Win32_VideoController")

$Displays = @(Get-DisplayTopology)
$TotalScanout = 0.0
foreach ($Display in $Displays) { $TotalScanout += $Display.ScanoutGigabytesPerSecond }

$PowerMode = $null
try { $PowerMode = Get-PowerModeOverlay } catch { Show-Output -ForegroundColor Yellow "Could not read the power mode: $($_.Exception.Message)" }
$SecurityStatus = Get-VirtualizationSecurityStatus
$ActiveScheme = (powercfg /getactivescheme) -join " "

# -----
# HWiNFO logging
# -----

$HWiNFOProcess = $null
$HWiNFOLog = "${OutputPath}\hwinfo_sensors.csv"
$HWiNFOExe = ""
if (-not $NoHWiNFO) {
    $HWiNFOExe = Find-HWiNFO -Path $HWiNFOPath
    if ($HWiNFOExe) {
        $HWiNFOProcess = Start-HWiNFOLogging -ExecutablePath $HWiNFOExe -LogPath $HWiNFOLog -PollRateMs ($Interval * 1000)
    } else {
        Show-Output "HWiNFO was not found. Install it with `"choco install hwinfo`" for detailed sensor data."
    }
}

# -----
# Sampling
# -----

Show-Output "Sampling performance counters for ${Duration} seconds at ${Interval} second intervals."
Show-Output "Reproduce the slowness now, e.g. by joining a video conference."

$CounterPaths = [ordered]@{
    CpuPercent = "\Processor Information(_Total)\% Processor Time"
    # Above 100 % means that the CPU is boosting above its nominal frequency.
    CpuPerformancePercent = "\Processor Information(_Total)\% Processor Performance"
    CpuMaxFrequencyPercent = "\Processor Information(_Total)\% of Maximum Frequency"
    CpuPrivilegedPercent = "\Processor Information(_Total)\% Privileged Time"
    CpuInterruptPercent = "\Processor Information(_Total)\% Interrupt Time"
    CpuDpcPercent = "\Processor Information(_Total)\% DPC Time"
    ProcessorQueueLength = "\System\Processor Queue Length"
    ContextSwitchesPerSec = "\System\Context Switches/sec"
    MemoryAvailableMB = "\Memory\Available MBytes"
    MemoryCommittedBytes = "\Memory\Committed Bytes"
    MemoryPagesPerSec = "\Memory\Pages/sec"
    DiskIdlePercent = "\PhysicalDisk(_Total)\% Idle Time"
    DiskReadLatencySec = "\PhysicalDisk(_Total)\Avg. Disk sec/Read"
    DiskWriteLatencySec = "\PhysicalDisk(_Total)\Avg. Disk sec/Write"
    DiskQueueLength = "\PhysicalDisk(_Total)\Current Disk Queue Length"
}
$Counters = Select-AvailableCounter -Counters $CounterPaths

# Map the normalized counter paths back to the column names
$KeyToColumn = @{}
foreach ($Name in $Counters.Keys) { $KeyToColumn[(ConvertTo-CounterKey $Counters[$Name])] = $Name }
$CounterList = @($Counters.Values)

$Samples = New-Object System.Collections.Generic.List[PSObject]
$SamplingStart = Get-Date
$ProcessesBefore = Get-ProcessSnapshot
$SamplingEnd = $SamplingStart.AddSeconds($Duration)
$SampleNumber = 0

while ((Get-Date) -lt $SamplingEnd) {
    $SampleNumber++
    $Row = [ordered]@{ Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") }
    foreach ($Name in $Counters.Keys) { $Row[$Name] = $null }

    if ($CounterList.Count -gt 0) {
        try {
            foreach ($Sample in (Get-Counter -Counter $CounterList -ErrorAction Stop).CounterSamples) {
                $Key = ConvertTo-CounterKey $Sample.Path
                if ($KeyToColumn.ContainsKey($Key)) {
                    $Row[$KeyToColumn[$Key]] = [math]::Round([double]$Sample.CookedValue, 3)
                }
            }
        } catch {
            Show-Output -ForegroundColor Yellow "Could not read the performance counters: $($_.Exception.Message)"
        }
    }

    $Gpu = Get-GpuEngineUsage
    $Row["GpuTotalPercent"] = $Gpu.Total
    $Row["Gpu3DPercent"] = $Gpu["3D"]
    $Row["GpuVideoDecodePercent"] = $Gpu.VideoDecode
    $Row["GpuVideoEncodePercent"] = $Gpu.VideoEncode
    $Row["GpuCopyPercent"] = $Gpu.Copy

    foreach ($Zone in (Get-ThermalZoneTemperature)) {
        $Row["Thermal_$($Zone.Name -replace '[^A-Za-z0-9]', '_')"] = $Zone.Celsius
    }

    $Samples.Add([PSCustomObject]$Row)

    # Show progress, since the sampling takes a while
    $Elapsed = ((Get-Date) - $SamplingStart).TotalSeconds
    Write-Progress `
        -Activity "Sampling performance counters" `
        -Status "Sample ${SampleNumber}, $([math]::Round($Elapsed)) / ${Duration} seconds" `
        -PercentComplete ([math]::Min(100, 100 * $Elapsed / $Duration))

    $SleepSeconds = $Interval - (((Get-Date) - $SamplingStart).TotalSeconds % $Interval)
    if ($SleepSeconds -gt 0 -and (Get-Date).AddSeconds($SleepSeconds) -lt $SamplingEnd) {
        Start-Sleep -Seconds $SleepSeconds
    } elseif ((Get-Date) -lt $SamplingEnd) {
        Start-Sleep -Milliseconds 200
    }
}
Write-Progress -Activity "Sampling performance counters" -Completed

$ProcessesAfter = Get-ProcessSnapshot
$ActualDuration = ((Get-Date) - $SamplingStart).TotalSeconds

if ($HWiNFOProcess) {
    Stop-HWiNFOLogging -Process $HWiNFOProcess
}

# -----
# Analysis
# -----

Show-Output "Analyzing the results."

$Samples | Export-Csv -Path $SamplesPath -NoTypeInformation -Encoding UTF8 -UseCulture

$TopProcesses = Get-TopProcessCpuUsage `
    -Before $ProcessesBefore `
    -After $ProcessesAfter `
    -ElapsedSeconds $ActualDuration
$TopProcesses | Export-Csv -Path $ProcessesPath -NoTypeInformation -Encoding UTF8 -UseCulture

function Get-ColumnStatistics {
    <#
    .SYNOPSIS
        Compute the statistics of one column of the samples.
    .PARAMETER Name
        The name of the column.
    #>
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory=$true)][string]$Name
    )
    $Values = New-Object System.Collections.Generic.List[double]
    foreach ($Sample in $Samples) {
        if ($Sample.PSObject.Properties.Name -contains $Name) {
            $Value = $Sample.$Name
            if ($null -ne $Value) { $Values.Add([double]$Value) }
        }
    }
    return Get-SampleStatistics -Values $Values.ToArray()
}

$ThrottleDuringRun = Get-ProcessorPowerEventSummary -StartTime $SamplingStart
$ThrottleHistory = Get-ProcessorPowerEventSummary -StartTime (Get-Date).AddDays(-$EventHistoryDays)
$EventsDuringRun = Get-EventLogSummary -StartTime $SamplingStart
$EventHistory = Get-EventLogSummary -StartTime (Get-Date).AddDays(-$EventHistoryDays) -MaxGroups 30

# -----
# Reporting
# -----

$Report = New-Object System.Collections.Generic.List[string]
function Add-Line {
    <#
    .SYNOPSIS
        Add a line to the summary report.
    .PARAMETER Text
        The text to add.
    #>
    param([string]$Text = "")
    $Report.Add($Text)
}

Add-Line "==============================================================================="
Add-Line " Performance diagnostics for ${env:ComputerName} at ${Timestamp}"
Add-Line "==============================================================================="
Add-Line ""
Add-Line "--- System ---"
Add-Line "Computer:        $($ComputerSystem.Manufacturer) $($ComputerSystem.Model) ($($ComputerSystem.SystemFamily))"
Add-Line "Operating system: $($OperatingSystem.Caption) $($OperatingSystem.Version)"
Add-Line "CPU:             $($Processor.Name)"
Add-Line "                 $($Processor.NumberOfCores) cores, $($Processor.NumberOfLogicalProcessors) logical processors, nominal $($Processor.MaxClockSpeed) MHz"
Add-Line "Memory:          $([math]::Round($ComputerSystem.TotalPhysicalMemory / 1GB, 1)) GB"
Add-Line "Power scheme:    ${ActiveScheme}"
if ($PowerMode) {
    Add-Line "Power mode:      $($PowerMode.Name)"
    Add-Line "                 On Lenovo ThinkPads this also selects the Intelligent Cooling mode."
}
Add-Line ""

Add-Line "--- Virtualization-based security ---"
Add-Line "Hypervisor present: $($SecurityStatus.HypervisorPresent)"
Add-Line "VBS status:         $($SecurityStatus.Status)"
Add-Line "Running:            $(@($SecurityStatus.Running) -join '; ')"
if (@($SecurityStatus.ConfiguredButNotRunning).Count -gt 0) {
    Add-Line "Configured but NOT running: $(@($SecurityStatus.ConfiguredButNotRunning) -join '; ')"
    Add-Line "  These features are configured but inactive. They do not add overhead of their own,"
    Add-Line "  but they also do not provide the protection they were enabled for."
}
Add-Line ""

Add-Line "--- Graphics adapters ---"
foreach ($Controller in $VideoControllers) {
    Add-Line "$($Controller.Name)"
    Add-Line "    Driver $($Controller.DriverVersion), $($Controller.DriverDate)"
}
Add-Line ""

Add-Line "--- Displays ---"
if (@($Displays).Count -eq 0) {
    Add-Line "No active displays were detected."
} else {
    foreach ($Display in $Displays) {
        $PrimaryTag = ""
        if ($Display.Primary) { $PrimaryTag = " [primary]" }
        # Built-in laptop panels usually do not report a model name in their EDID.
        $Name = $Display.MonitorName
        if (-not $Name -and $Display.ConnectionType -eq "Internal") { $Name = "built-in display" }
        if (-not $Name) { $Name = "unknown model" }
        $Label = "$($Display.Manufacturer) ${Name}".Trim()
        Add-Line "$($Display.DeviceName)${PrimaryTag}: ${Label}"
        Add-Line "    $($Display.Width)x$($Display.Height) @ $($Display.RefreshRate) Hz over $($Display.ConnectionType)"
        Add-Line "    Scanout bandwidth: $($Display.ScanoutGigabytesPerSecond) GB/s"
    }
    Add-Line ""
    Add-Line "Total scanout bandwidth: $([math]::Round($TotalScanout, 2)) GB/s"
    Add-Line "  This is the memory bandwidth consumed by merely displaying the desktop,"
    Add-Line "  excluding all rendering and compositing. On a system with an integrated GPU"
    Add-Line "  this competes with the CPU for the same system memory bandwidth."
    $RefreshRates = @($Displays | ForEach-Object { $_.RefreshRate } | Sort-Object -Unique)
    if ($RefreshRates.Count -gt 1) {
        Add-Line "  Note: the displays run at different refresh rates ($($RefreshRates -join ', ') Hz)."
        Add-Line "  Mixed refresh rates can prevent the GPU from entering its low power memory states."
    }
}
Add-Line ""

Add-Line "--- CPU frequency and load over $([math]::Round($ActualDuration)) seconds ($(@($Samples).Count) samples) ---"
$MetricDescriptions = [ordered]@{
    CpuPercent = "CPU utilization (%)"
    CpuPerformancePercent = "CPU performance vs nominal (%)"
    CpuMaxFrequencyPercent = "CPU frequency vs maximum (%)"
    CpuPrivilegedPercent = "Kernel mode time (%)"
    CpuInterruptPercent = "Interrupt time (%)"
    CpuDpcPercent = "DPC time (%)"
    ProcessorQueueLength = "Processor queue length"
    MemoryAvailableMB = "Available memory (MB)"
    DiskIdlePercent = "Disk idle time (%)"
    DiskReadLatencySec = "Disk read latency (s)"
    DiskWriteLatencySec = "Disk write latency (s)"
    GpuTotalPercent = "GPU utilization, all engines (%)"
    Gpu3DPercent = "GPU 3D engine (%)"
    GpuVideoDecodePercent = "GPU video decode engine (%)"
    GpuVideoEncodePercent = "GPU video encode engine (%)"
}
Add-Line ("{0,-38} {1,9} {2,9} {3,9} {4,9}" -f "Metric", "Min", "Median", "P95", "Max")
$Statistics = @{}
foreach ($Metric in $MetricDescriptions.Keys) {
    $Stats = Get-ColumnStatistics -Name $Metric
    $Statistics[$Metric] = $Stats
    if ($Stats.Count -eq 0) { continue }
    Add-Line ("{0,-38} {1,9} {2,9} {3,9} {4,9}" -f $MetricDescriptions[$Metric], $Stats.Min, $Stats.Median, $Stats.P95, $Stats.Max)
}
Add-Line ""

Add-Line "--- Top processes by CPU time during the measurement ---"
Add-Line ("{0,-32} {1,8} {2,12} {3,10} {4,10}" -f "Process", "PID", "CPU seconds", "% of core", "Memory MB")
foreach ($Process in $TopProcesses) {
    Add-Line ("{0,-32} {1,8} {2,12} {3,10} {4,10}" -f $Process.Name, $Process.Id, $Process.CpuSeconds, $Process.CpuPercent, $Process.WorkingSetMB)
}
Add-Line ""

Add-Line "--- Firmware CPU throttling (Kernel-Processor-Power event 37) ---"
Add-Line "This event means that the system firmware, not Windows, is limiting the CPU speed."
Add-Line "It indicates thermal or power limit throttling enforced by the embedded controller."
Add-Line "The reported seconds are counted per processor since the previous report for that"
Add-Line "processor, so the figures below are for the worst affected processor."
Add-Line ""
Add-Line "During the measurement: $($ThrottleDuringRun.EventCount) events, worst processor throttled for $($ThrottleDuringRun.WorstProcessorSeconds) s"
Add-Line "Last ${EventHistoryDays} days:      $($ThrottleHistory.EventCount) events, worst processor throttled for $($ThrottleHistory.WorstProcessorSeconds) s"
if ($ThrottleHistory.EventCount -gt 0) {
    $LongestHours = [math]::Round($ThrottleHistory.LongestSingleReportSeconds / 3600.0, 1)
    Add-Line "Longest single report: $($ThrottleHistory.LongestSingleReportSeconds) s (${LongestHours} h) of continuous throttling"
    Add-Line "Affected processors:   $($ThrottleHistory.AffectedProcessorCount) of $($Processor.NumberOfLogicalProcessors) ($($ThrottleHistory.AffectedProcessors))"
    Add-Line "Most recent event:     $($ThrottleHistory.LastEvent)"
}
Add-Line ""

# -----
# Findings
# -----

$Findings = New-Object System.Collections.Generic.List[string]

if ($ThrottleDuringRun.EventCount -gt 0) {
    $Findings.Add("The system firmware throttled the CPU during the measurement " +
        "($($ThrottleDuringRun.WorstProcessorSeconds) s on the worst affected processor). " +
        "This is thermal or power limit throttling, not a Windows setting.")
}
# Long reports are ambiguous, so they are reported without drawing a conclusion.
# The counter is not reset by sleep, so a report spanning a whole day usually just means
# that the machine was asleep or idle, not that it was throttling under load.
if ($ThrottleHistory.LongestSingleReportSeconds -ge 3600) {
    $Hours = [math]::Round($ThrottleHistory.LongestSingleReportSeconds / 3600.0, 1)
    $Findings.Add("The longest single firmware throttling report over the last ${EventHistoryDays} days " +
        "covers ${Hours} hours. Note that this counter keeps running while the machine sleeps or idles, " +
        "so a report spanning a whole day is not by itself evidence of throttling under load. " +
        "Only events whose duration is short compared to their reporting interval, and events that " +
        "occur while the machine is busy, indicate real throttling.")
}
if ($Statistics.ContainsKey("CpuMaxFrequencyPercent") -and $Statistics["CpuMaxFrequencyPercent"].Count -gt 0) {
    $Median = $Statistics["CpuMaxFrequencyPercent"].Median
    if ($Median -lt 80) {
        $Findings.Add("The CPU ran at only ${Median} % of its maximum frequency (median). " +
            "If the CPU was not idle, this suggests it is power or thermally limited.")
    }
}
if ($Statistics.ContainsKey("CpuDpcPercent") -and $Statistics["CpuDpcPercent"].Count -gt 0) {
    $Dpc = $Statistics["CpuDpcPercent"].P95
    $Interrupt = 0
    if ($Statistics.ContainsKey("CpuInterruptPercent") -and $Statistics["CpuInterruptPercent"].Count -gt 0) {
        $Interrupt = $Statistics["CpuInterruptPercent"].P95
    }
    if (($Dpc + $Interrupt) -gt 10) {
        $Findings.Add("High interrupt and DPC time (P95: ${Dpc} % DPC, ${Interrupt} % interrupt). " +
            "This points to a driver problem, often a network, storage, USB or graphics driver.")
    }
}
if ($Statistics.ContainsKey("DiskReadLatencySec") -and $Statistics["DiskReadLatencySec"].Count -gt 0) {
    $Latency = $Statistics["DiskReadLatencySec"].P95
    if ($Latency -gt 0.025) {
        $Findings.Add("High disk read latency (P95: $([math]::Round($Latency * 1000)) ms). " +
            "A healthy NVMe SSD should stay well below 10 ms.")
    }
}
if ($Statistics.ContainsKey("MemoryAvailableMB") -and $Statistics["MemoryAvailableMB"].Count -gt 0) {
    $AvailableMB = $Statistics["MemoryAvailableMB"].Min
    if ($AvailableMB -lt 1024) {
        $Findings.Add("Available memory dropped to ${AvailableMB} MB. The system is likely paging.")
    }
}
if ($TotalScanout -gt 4) {
    $IntegratedOnly = -not (@($VideoControllers | Where-Object { $_.Name -notmatch "Intel|AMD Radeon\(TM\) Graphics|Microsoft" }).Count -gt 0)
    $Extra = ""
    if ($IntegratedOnly) { $Extra = " All of it is handled by the integrated GPU, which shares memory bandwidth with the CPU." }
    $Findings.Add("The displays require $([math]::Round($TotalScanout, 2)) GB/s of scanout bandwidth.${Extra} " +
        "Lowering the refresh rate or the resolution of the external displays reduces this.")
}
if (@($SecurityStatus.ConfiguredButNotRunning).Count -gt 0) {
    $Findings.Add("These security features are configured but not running: " +
        "$(@($SecurityStatus.ConfiguredButNotRunning) -join '; '). Check whether they are licensed and supported.")
}
if ($PowerMode -and $PowerMode.Name -ne "Best performance") {
    $Findings.Add("The Windows power mode is `"$($PowerMode.Name)`". On a Lenovo ThinkPad this also " +
        "selects the Intelligent Cooling mode, which limits the sustained CPU power. " +
        "Use Set-PowerMode.ps1 to change it.")
}

Add-Line "--- Findings ---"
if ($Findings.Count -eq 0) {
    Add-Line "No significant performance problems were detected during the measurement."
} else {
    $Index = 1
    foreach ($Finding in $Findings) {
        Add-Line "${Index}. ${Finding}"
        $Index++
    }
}
Add-Line ""

# -----
# Event log details
# -----

$EventReport = New-Object System.Collections.Generic.List[string]
$EventReport.Add("--- Event log warnings and errors during the measurement ---")
if (@($EventsDuringRun).Count -eq 0) {
    $EventReport.Add("None.")
} else {
    foreach ($Event in $EventsDuringRun) {
        $EventReport.Add("[$($Event.Count)x] $($Event.Provider) $($Event.Id) ($($Event.Level))")
        $EventReport.Add("    $($Event.Message)")
    }
}
$EventReport.Add("")
$EventReport.Add("--- Event log warnings and errors over the last ${EventHistoryDays} days ---")
if (@($EventHistory).Count -eq 0) {
    $EventReport.Add("None.")
} else {
    foreach ($Event in $EventHistory) {
        $EventReport.Add("[$($Event.Count)x] $($Event.Provider) $($Event.Id) ($($Event.Level))")
        $EventReport.Add("    $($Event.Message)")
    }
}
$EventReport -join "`r`n" | Set-Content -Path $EventsPath -Encoding UTF8

# -----
# Output
# -----

$ReportText = $Report -join "`r`n"
$ReportText | Set-Content -Path $SummaryPath -Encoding UTF8
Show-Output ""
Show-Output $ReportText
Show-Output ""
Show-Output -ForegroundColor Green "The performance diagnostics are ready."
Show-Output -ForegroundColor Green "Summary:        ${SummaryPath}"
Show-Output -ForegroundColor Green "Samples:        ${SamplesPath}"
Show-Output -ForegroundColor Green "Processes:      ${ProcessesPath}"
Show-Output -ForegroundColor Green "Event log:      ${EventsPath}"
if ($HWiNFOProcess -and (Test-Path -LiteralPath $HWiNFOLog)) {
    Show-Output -ForegroundColor Green "HWiNFO sensors: ${HWiNFOLog}"
}
