<#
.SYNOPSIS
    Lint the PowerShell scripts of this repository with PSScriptAnalyzer
.DESCRIPTION
    Installs PSScriptAnalyzer for the current user if it is missing, analyzes every
    *.ps1 and *.psm1 file of the repository with the rules configured in
    PSScriptAnalyzerSettings.psd1, prints a summary, and exits with a non-zero exit code
    if the failure threshold is reached.

    This script is intentionally standalone. It does not dot-source Utils.ps1, so that it
    can also be run on a non-Windows CI runner with PowerShell 7.
.PARAMETER Path
    The file or directory to analyze. Defaults to the directory of this script.
.PARAMETER Severity
    The severities to report. Defaults to all of them.
.PARAMETER FailOn
    The lowest severity that makes the script exit with a non-zero exit code.
    "None" never fails. Defaults to "Error".
.PARAMETER Fix
    Apply the automatic formatting fixes of Invoke-Formatter to the analyzed files
    before reporting. The files are rewritten in place, preserving their encoding.
.EXAMPLE
    pwsh ./Invoke-Lint.ps1
.EXAMPLE
    pwsh ./Invoke-Lint.ps1 -Fix
.EXAMPLE
    pwsh ./Invoke-Lint.ps1 -Severity Error, Warning -FailOn Warning
#>

[CmdletBinding()]
param(
    [string]$Path = $PSScriptRoot,
    [ValidateSet("Error", "Warning", "Information")]
    [string[]]$Severity = @("Error", "Warning", "Information"),
    [ValidateSet("Error", "Warning", "Information", "None")]
    [string]$FailOn = "Error",
    [switch]$Fix
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

# Directories that are never linted. These are either ignored by Git (see .gitignore)
# or contain no code of our own.
$ExcludedDirectories = @(
    ".git",
    ".idea",
    "downloads",
    "lib",
    "logs",
    "reports",
    "venv"
)
$LintExtensions = @(".ps1", ".psm1")
$SettingsPath = Join-Path -Path $PSScriptRoot -ChildPath "PSScriptAnalyzerSettings.psd1"

# The formatting rules are applied as fixes instead of being reported, so they are
# configured here instead of in PSScriptAnalyzerSettings.psd1. The values match the
# existing style of the repository: 4-space indentation and the one-true-brace style.
$FormatterSettings = @{
    IncludeRules = @(
        "PSPlaceOpenBrace",
        "PSPlaceCloseBrace",
        "PSUseConsistentIndentation",
        "PSUseConsistentWhitespace",
        "PSUseCorrectCasing"
    )
    Rules        = @{
        PSPlaceOpenBrace           = @{
            Enable             = $true
            OnSameLine         = $true
            NewLineAfter       = $true
            IgnoreOneLineBlock = $true
        }
        PSPlaceCloseBrace          = @{
            Enable             = $true
            NewLineAfter       = $false
            IgnoreOneLineBlock = $true
            NoEmptyLineBefore  = $false
        }
        PSUseConsistentIndentation = @{
            Enable              = $true
            IndentationSize     = 4
            Kind                = "space"
            PipelineIndentation = "IncreaseIndentationForFirstPipeline"
        }
        PSUseConsistentWhitespace  = @{
            Enable = $true
        }
        PSUseCorrectCasing         = @{
            Enable = $true
        }
    }
}

function Install-Analyzer {
    <#
    .SYNOPSIS
        Make sure that the PSScriptAnalyzer module is available
    #>
    [OutputType([void])]
    param()
    if (Get-Module -ListAvailable -Name "PSScriptAnalyzer") {
        Import-Module -Name "PSScriptAnalyzer"
        return
    }
    Write-Output "PSScriptAnalyzer was not found. Installing it for the current user."
    if (Get-Command -Name "Install-PSResource" -ErrorAction SilentlyContinue) {
        Install-PSResource -Name "PSScriptAnalyzer" -Scope CurrentUser -TrustRepository
    } else {
        Install-Module -Name "PSScriptAnalyzer" -Scope CurrentUser -Force
    }
    Import-Module -Name "PSScriptAnalyzer"
}

function Add-LintTarget {
    <#
    .SYNOPSIS
        Collect the paths of the files to lint from a directory tree
    .PARAMETER Directory
        The directory to walk through
    .PARAMETER Accumulator
        The list the found paths are added to
    #>
    [OutputType([void])]
    param(
        [Parameter(Mandatory = $true)][string]$Directory,
        # AllowEmptyCollection is required, since a mandatory parameter otherwise rejects
        # the empty list that the first call passes in.
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]]$Accumulator
    )
    foreach ($Item in @(Get-ChildItem -LiteralPath "${Directory}" -Force)) {
        if ($Item.PSIsContainer) {
            if ($ExcludedDirectories -notcontains $Item.Name) {
                Add-LintTarget -Directory $Item.FullName -Accumulator $Accumulator
            }
        } elseif ($LintExtensions -contains $Item.Extension) {
            $Accumulator.Add($Item.FullName)
        }
    }
}

function Get-LintTarget {
    <#
    .SYNOPSIS
        Resolve the -Path parameter into the list of files to lint
    .PARAMETER Target
        The file or directory given on the command line
    #>
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $true)][string]$Target
    )
    $Resolved = (Resolve-Path -LiteralPath "${Target}").ProviderPath
    $Files = [System.Collections.Generic.List[string]]::new()
    if (Test-Path -LiteralPath "${Resolved}" -PathType Container) {
        Add-LintTarget -Directory "${Resolved}" -Accumulator $Files
    } else {
        $Files.Add($Resolved)
    }
    $Files.Sort()
    return $Files.ToArray()
}

function Repair-File {
    <#
    .SYNOPSIS
        Apply the automatic formatting fixes to a single file
    .DESCRIPTION
        Returns $true if the file was changed. The original encoding, with or without a
        byte order mark, is preserved. [IO.File] is used instead of Get-Content and
        Set-Content, because the meaning of their -Encoding values differs between
        Windows PowerShell 5.1 and PowerShell 7.
    .PARAMETER FilePath
        The file to format
    #>
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)][string]$FilePath
    )
    $Bytes = [IO.File]::ReadAllBytes($FilePath)
    $HasBom = ($Bytes.Length -ge 3) -and ($Bytes[0] -eq 0xEF) -and ($Bytes[1] -eq 0xBB) -and ($Bytes[2] -eq 0xBF)
    $Original = [IO.File]::ReadAllText($FilePath)
    $Formatted = Invoke-Formatter -ScriptDefinition $Original -Settings $FormatterSettings
    if ($Formatted -ceq $Original) {
        return $false
    }
    [IO.File]::WriteAllText($FilePath, $Formatted, [Text.UTF8Encoding]::new($HasBom))
    return $true
}

function Write-Summary {
    <#
    .SYNOPSIS
        Print the findings grouped by file, and the counts by severity
    .PARAMETER Finding
        The diagnostic records returned by PSScriptAnalyzer
    .PARAMETER Root
        The directory the reported paths are made relative to
    #>
    [OutputType([void])]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Finding,
        [Parameter(Mandatory = $true)][string]$Root
    )
    $Prefix = $Root.TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    foreach ($Group in ($Finding | Group-Object -Property "ScriptPath" | Sort-Object -Property "Name")) {
        $Relative = $Group.Name
        if ($Relative.StartsWith("${Prefix}")) {
            $Relative = $Relative.Substring($Prefix.Length).TrimStart([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
        }
        Write-Output ""
        Write-Output "${Relative} ($($Group.Count))"
        foreach ($Record in ($Group.Group | Sort-Object -Property "Line", "Column")) {
            Write-Output ("  {0,4}:{1,-4} {2,-11} {3,-45} {4}" -f `
                    $Record.Line, $Record.Column, $Record.Severity, $Record.RuleName, $Record.Message)
        }
    }

    Write-Output ""
    Write-Output "Findings by severity:"
    foreach ($Name in @("Error", "Warning", "Information")) {
        $Count = @($Finding | Where-Object { "$($_.Severity)" -eq $Name }).Count
        Write-Output ("  {0,-12} {1}" -f $Name, $Count)
    }
    Write-Output ("  {0,-12} {1}" -f "Total", $Finding.Count)
}

# -----
# Main
# -----
if (-not (Test-Path -LiteralPath "${SettingsPath}")) {
    throw "The settings file `"${SettingsPath}`" was not found."
}
Install-Analyzer

$Root = (Resolve-Path -LiteralPath "${Path}").ProviderPath
$Targets = Get-LintTarget -Target "${Path}"
if ($Targets.Count -eq 0) {
    Write-Output "No PowerShell files were found in `"${Root}`"."
    exit 0
}
Write-Output "Analyzing $($Targets.Count) file(s) in `"${Root}`" with `"${SettingsPath}`"."

if ($Fix) {
    $Repaired = 0
    foreach ($Target in $Targets) {
        if (Repair-File -FilePath "${Target}") {
            $Repaired++
            Write-Output "Formatted: ${Target}"
        }
    }
    Write-Output "Applied formatting fixes to ${Repaired} file(s)."
}

$Findings = [System.Collections.Generic.List[object]]::new()
$AnalyzerErrors = [System.Collections.Generic.List[string]]::new()
foreach ($Target in $Targets) {
    $RunErrors = @()
    $Records = @(Invoke-ScriptAnalyzer `
            -Path "${Target}" `
            -Settings "${SettingsPath}" `
            -Severity $Severity `
            -ErrorAction SilentlyContinue `
            -ErrorVariable RunErrors)
    foreach ($Record in $Records) {
        $Findings.Add($Record)
    }
    foreach ($RunError in $RunErrors) {
        $AnalyzerErrors.Add("${Target}: $($RunError.Exception.Message)")
    }
}

Write-Summary -Finding $Findings.ToArray() -Root "${Root}"

if ($AnalyzerErrors.Count -gt 0) {
    # These are problems in the analyzer input itself, for example a
    # SuppressMessageAttribute that no longer matches any finding. They are reported for
    # visibility, but they do not affect the exit code.
    Write-Output ""
    Write-Output "PSScriptAnalyzer reported $($AnalyzerErrors.Count) problem(s) while analyzing:"
    foreach ($AnalyzerError in $AnalyzerErrors) {
        Write-Output "  ${AnalyzerError}"
    }
}

$FailSeverities = switch ($FailOn) {
    "Error" { @("Error") }
    "Warning" { @("Error", "Warning") }
    "Information" { @("Error", "Warning", "Information") }
    default { @() }
}
$Fatal = @($Findings | Where-Object { $FailSeverities -contains "$($_.Severity)" })
Write-Output ""
if ($Fatal.Count -gt 0) {
    Write-Output "Linting failed: $($Fatal.Count) finding(s) of severity ${FailOn} or higher."
    exit 1
}
Write-Output "Linting passed. The failure threshold is `"${FailOn}`"."
exit 0
