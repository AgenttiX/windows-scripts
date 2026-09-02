<#
.SYNOPSIS
    Update group policies and see that they are applied properly.
#>

Set-StrictMode -Version 3.0

gpupdate /force
gpresult /h "report.html" /f
.\report.html
