<#
.SYNOPSIS
    Update group policies and see that they are applied properly.
#>

gpupdate /force
gpresult /h "report.html" /f
.\report.html
