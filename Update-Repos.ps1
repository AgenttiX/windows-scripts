<#
.SYNOPSIS
    Refresh Git repositories
#>

$Repos = Get-ChildItem -Directory ((Get-Item "${PSScriptRoot}").Parent.FullName)

foreach ($Repo in $Repos) {
    Set-Location $Repo.FullName
    git pull
}
