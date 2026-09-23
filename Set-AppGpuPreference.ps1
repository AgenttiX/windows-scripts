<#
.SYNOPSIS
    Show or set which GPU an application uses on a hybrid graphics laptop.
.DESCRIPTION
    On a laptop with both an integrated and a discrete GPU, Windows decides per application
    which one to use. The choice is stored per executable in the registry, and is the same
    setting as the one in Settings -> System -> Display -> Graphics.

    This is useful on machines where the integrated GPU is the bottleneck. Video conferencing
    loads the integrated GPU with video encoding, video decoding and compositing, while it is
    at the same time driving the displays and sharing its memory bandwidth with the CPU.
    Moving such an application to the discrete GPU can relieve the integrated one.

    This is a trade-off rather than a certain improvement, and it is worth measuring:

    - The discrete GPU has its own memory, so it relieves the memory bandwidth of the
      integrated GPU and of the CPU.
    - On a hybrid graphics laptop the displays are still driven by the integrated GPU,
      so every rendered frame has to be copied back to it over PCIe.
    - Waking the discrete GPU adds power draw, which on a laptop with a shared power and
      thermal budget can leave less headroom for the CPU.

    Measure the result with Get-PerformanceDiagnostics.ps1 before and after the change.
.PARAMETER Preference
    The GPU to use. "HighPerformance" is normally the discrete GPU and "PowerSaving" the
    integrated one. "Auto" removes the setting and lets Windows decide.
.PARAMETER Application
    A predefined set of executables to configure. "Teams" covers both the Teams executable
    and the WebView2 runtime that actually renders and encodes its video.
.PARAMETER Path
    The full paths of additional executables to configure.
.PARAMETER List
    Show the current settings instead of changing them.
.EXAMPLE
    .\Set-AppGpuPreference.ps1 -List
    Show which applications have a GPU preference configured.
.EXAMPLE
    .\Set-AppGpuPreference.ps1 -Application Teams -Preference HighPerformance
    Move Microsoft Teams and its WebView2 runtime to the discrete GPU.
.EXAMPLE
    .\Set-AppGpuPreference.ps1 -Application Teams -Preference Auto
    Undo the change and let Windows decide again.
.NOTES
    The applications have to be restarted for the change to take effect. For Teams this
    means quitting it from the system tray, not just closing the window.

    The setting is per user and is stored under HKCU, so it does not require
    administrator privileges.
#>

param(
    [ValidateSet("HighPerformance", "PowerSaving", "Auto")]
    [string]$Preference = "HighPerformance",
    [ValidateSet("Teams")]
    [string[]]$Application,
    [string[]]$Path,
    [switch]$List
)

Set-StrictMode -Version 3.0
. "${PSScriptRoot}\Utils.ps1"

$RegistryPath = "HKCU:\Software\Microsoft\DirectX\UserGpuPreferences"

# The values that Windows itself writes for the options of the Graphics settings page
$PreferenceValues = @{
    "Auto" = 0
    "PowerSaving" = 1
    "HighPerformance" = 2
}
$PreferenceNames = @{
    "0" = "Let Windows decide"
    "1" = "Power saving (usually the integrated GPU)"
    "2" = "High performance (usually the discrete GPU)"
}

function Get-TeamsExecutable {
    <#
    .SYNOPSIS
        Get the executables that render and encode the video of Microsoft Teams.
    .DESCRIPTION
        The new Teams client is a packaged application whose install path contains its
        version number, so the path changes at every update and has to be resolved
        dynamically. The actual rendering and video encoding happens in the WebView2
        runtime, which is a separate executable and needs the same setting.
    #>
    [OutputType([System.Array])]
    param()
    $Executables = New-Object System.Collections.Generic.List[string]

    $Package = Get-AppxPackage -Name "MSTeams" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($Package) {
        $TeamsExe = Join-Path -Path $Package.InstallLocation -ChildPath "ms-teams.exe"
        if (Test-Path -LiteralPath $TeamsExe) { $Executables.Add($TeamsExe) }
    } else {
        Show-Information -ForegroundColor Yellow "The Microsoft Teams package was not found."
    }

    # The WebView2 runtime keeps one directory per installed version.
    # Configure all of them, since Teams may switch between them at an update.
    $WebViewRoots = @(
        "${env:ProgramFiles(x86)}\Microsoft\EdgeWebView\Application",
        "${env:ProgramFiles}\Microsoft\EdgeWebView\Application"
    )
    foreach ($Root in $WebViewRoots) {
        if (-not $Root -or -not (Test-Path -LiteralPath $Root)) { continue }
        foreach ($Directory in (Get-ChildItem -LiteralPath $Root -Directory -ErrorAction SilentlyContinue)) {
            $WebViewExe = Join-Path -Path $Directory.FullName -ChildPath "msedgewebview2.exe"
            if (Test-Path -LiteralPath $WebViewExe) { $Executables.Add($WebViewExe) }
        }
    }
    if ($Executables.Count -eq 0) {
        Show-Information -ForegroundColor Yellow "No Microsoft Teams executables were found."
    }
    return $Executables.ToArray()
}

function Show-GpuPreference {
    <#
    .SYNOPSIS
        Show the currently configured GPU preferences.
    .PARAMETER Filter
        An optional regular expression to limit which applications are shown.
    #>
    [OutputType([void])]
    param(
        [string]$Filter
    )
    if (-not (Test-Path -LiteralPath $RegistryPath)) {
        Show-Output "No GPU preferences have been configured."
        return
    }
    $Key = Get-Item -LiteralPath $RegistryPath
    $Names = @($Key.GetValueNames() | Where-Object { $_ })
    if ($Filter) { $Names = @($Names | Where-Object { $_ -match $Filter }) }
    if ($Names.Count -eq 0) {
        Show-Output "No matching GPU preferences have been configured."
        return
    }
    foreach ($Name in ($Names | Sort-Object)) {
        $Value = Get-ItemPropertyValue -LiteralPath $RegistryPath -Name $Name
        $Description = "unrecognized value `"${Value}`""
        if ($Value -match "GpuPreference=(\d+)") {
            $Number = $Matches[1]
            if ($PreferenceNames.ContainsKey($Number)) { $Description = $PreferenceNames[$Number] }
            else { $Description = "unknown preference ${Number}" }
        }
        Show-Output "${Name}"
        Show-Output "    ${Description}"
    }
}

function Set-GpuPreference {
    <#
    .SYNOPSIS
        Set the GPU preference of a single executable.
    .PARAMETER ExecutablePath
        The full path of the executable.
    .PARAMETER PreferenceName
        The name of the preference to apply.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        "PSUseShouldProcessForStateChangingFunctions",
        "",
        Justification="Interactive administration script"
    )]
    [OutputType([void])]
    param(
        [Parameter(Mandatory=$true)][string]$ExecutablePath,
        [Parameter(Mandatory=$true)][string]$PreferenceName
    )
    if (-not (Test-Path -LiteralPath $RegistryPath)) {
        New-Item -Path $RegistryPath -Force | Out-Null
    }
    if ($PreferenceName -eq "Auto") {
        # Windows treats a missing value as "let Windows decide",
        # which is cleaner than storing GpuPreference=0.
        Remove-ItemProperty -LiteralPath $RegistryPath -Name $ExecutablePath -ErrorAction SilentlyContinue
        Show-Output "Removed the setting of `"${ExecutablePath}`"."
        return
    }
    $Value = "GpuPreference=$($PreferenceValues[$PreferenceName]);"
    Set-ItemProperty -LiteralPath $RegistryPath -Name $ExecutablePath -Value $Value -Type String
    Show-Output "Set `"${ExecutablePath}`" to ${PreferenceName}."
}

# -----
# Main
# -----

Show-Output "Configuring the GPU preferences of applications."

$Adapters = @(Get-CimInstance -ClassName "Win32_VideoController" | Select-Object -ExpandProperty Name)
if ($Adapters.Count -lt 2) {
    Show-Output -ForegroundColor Yellow (
        "Only one graphics adapter was detected ($($Adapters -join ', ')). " +
        "The GPU preference has an effect only on hybrid graphics systems."
    )
} else {
    Show-Output "Graphics adapters: $($Adapters -join ', ')"
}
Show-Output ""

if ($List) {
    Show-GpuPreference
    exit
}

$Targets = New-Object System.Collections.Generic.List[string]
foreach ($Name in @($Application)) {
    if ($Name -eq "Teams") {
        foreach ($Executable in (Get-TeamsExecutable)) { $Targets.Add($Executable) }
    }
}
foreach ($Item in @($Path)) {
    if (-not $Item) { continue }
    if (Test-Path -LiteralPath $Item) {
        $Targets.Add((Resolve-Path -LiteralPath $Item).Path)
    } else {
        Show-Output -ForegroundColor Yellow "The executable `"${Item}`" was not found. Skipping it."
    }
}

if ($Targets.Count -eq 0) {
    Show-Output -ForegroundColor Red "No applications were given. Use -Application, -Path or -List."
    exit 1
}

foreach ($Target in ($Targets | Sort-Object -Unique)) {
    Set-GpuPreference -ExecutablePath $Target -PreferenceName $Preference
}

Show-Output ""
Show-Output -ForegroundColor Green "The GPU preferences have been updated."
Show-Output -ForegroundColor Green (
    "Restart the applications for the change to take effect. " +
    "Teams has to be quit from the system tray, not just closed."
)
Show-Output ""
Show-Output "Verify the result by running Get-PerformanceDiagnostics.ps1 during a video conference"
Show-Output "and comparing the GPU engine utilization and the CPU throttling before and after."
