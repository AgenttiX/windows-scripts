<#
.SYNOPSIS
    Show or set the Windows power mode.
.DESCRIPTION
    On Windows 10 and 11 the power mode is implemented as an overlay on top of the active
    power scheme. It is therefore not visible in the Control Panel power options, which is
    why modern laptops often appear to have only the "Balanced" scheme available.

    On Lenovo ThinkPads this overlay also selects the Intelligent Cooling mode, which
    controls the sustained power and thermal limits enforced by the firmware. On these
    machines the Fn+Q shortcut does not exist, and the power mode is the supported way
    to switch between the cooling modes.

    Changing the power mode does not require administrator privileges.
.PARAMETER Mode
    The power mode to activate. If omitted, the current power mode is shown.
.EXAMPLE
    .\Set-PowerMode.ps1
    Show the current power mode.
.EXAMPLE
    .\Set-PowerMode.ps1 -Mode BestPerformance
    Switch to the best performance mode, which on a ThinkPad also selects the
    Performance cooling mode.
.LINK
    https://learn.microsoft.com/en-us/windows/win32/api/powersetting/nf-powersetting-powersetactiveoverlayscheme
#>

param(
    [ValidateSet("BestPerformance", "Balanced", "BestPowerEfficiency")]
    [string]$Mode
)

Set-StrictMode -Version 3.0
. "${PSScriptRoot}\Utils.ps1"

$Current = Get-PowerModeOverlay

if (-not $Mode) {
    Show-Output "The current power mode is `"$($Current.Name)`" ($($Current.Guid))."
    Show-Output ""
    Show-Output "Set it with one of:"
    Show-Output "    .\Set-PowerMode.ps1 -Mode BestPerformance"
    Show-Output "    .\Set-PowerMode.ps1 -Mode Balanced"
    Show-Output "    .\Set-PowerMode.ps1 -Mode BestPowerEfficiency"
    exit
}

Show-Output "Changing the power mode from `"$($Current.Name)`"."
Set-PowerModeOverlay -Mode $Mode

$New = Get-PowerModeOverlay
if ($New.Guid -eq $Current.Guid -and $Current.Name -ne (Get-PowerModeOverlayName ([Guid]$New.Guid))) {
    Show-Output -ForegroundColor Red "The power mode did not change."
    exit 1
}
Show-Output -ForegroundColor Green "The power mode is now `"$($New.Name)`"."

if ($Mode -eq "BestPerformance") {
    Show-Output ""
    Show-Output "On a Lenovo ThinkPad this also selects the Performance cooling mode."
    Show-Output "Note that the power mode is reset by some Lenovo Vantage updates and by"
    Show-Output "switching between battery and mains power, so verify it if the machine feels slow."
}
