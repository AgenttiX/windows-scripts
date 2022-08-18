<#
.SYNOPSIS
    Adds an Authenticode signature to a PowerShell script or other file.
.DESCRIPTION
    When the PowerShell execution policy is set to RemoteSigned or higher,
    executing remote scripts requires that they have a valid signature.
    To sign scripts you need to have a code signing certificate that is trusted by the client machines.
    This certificate can be generated with e.g. Active Directory Certificate Services.
.LINK
    https://adamtheautomator.com/how-to-sign-powershell-script/
.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-authenticodesignature
#>
param(
    [Parameter(Position=0, Mandatory=$true)][string]$FilePath,
    [string]$CertPath,
    [string]$TimestampServer = "http://timestamp.digicert.com"
)

# Get the code-signing certificate from the local computer's certificate store
if ($PSBoundParameters.ContainsKey("CertPath")) {
    $Certificate = Get-item "${CertPath}"
} else {
    $Certificate = Get-ChildItem "Cert:\CurrentUser\My" | Where-Object {$_.EnhancedKeyUsageList.FriendlyName -eq "Code Signing"}
}
if ($null -eq $Certificate) {
    Write-Output "Could not sign, as no signing certificate was found."
    return
}

# Adding a timestamp ensures that your code will not expire when the signing certificate expires.
Set-AuthenticodeSignature -FilePath "${FilePath}" -Certificate $Certificate -HashAlgorithm "SHA256" -TimestampServer "${TimestampServer}"
