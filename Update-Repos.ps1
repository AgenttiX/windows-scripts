<#
.SYNOPSIS
    Refresh Git repositories
#>

Set-StrictMode -Version 3.0
$Repos = Get-ChildItem -Directory ((Get-Item "${PSScriptRoot}").Parent.FullName)

foreach ($Repo in $Repos) {
    Set-Location $Repo.FullName
    git pull
}
