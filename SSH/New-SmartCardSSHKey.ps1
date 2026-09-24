<#
.SYNOPSIS
    Create an SSH key on a TPM virtual smart card
.DESCRIPTION
    Creates a non-exportable RSA key and a self-signed certificate for it on the virtual smart card
    created by New-TPMVirtualSmartCard.ps1, and saves the public key in the OpenSSH format.
    Alternatively, an existing certificate can be used with -Thumbprint.

    The key is used through the Pageant of PuTTY-CAC, so PuTTY-CAC has to be installed first.
    It can be installed with Install-Software.ps1, or with either winget or Chocolatey:
    winget install --exact --id NoMoreFood.PuTTY-CAC
    choco install putty-cac

    Run this as a normal user, since the certificate is stored in the certificate store of the current user.
    Windows asks for the PIN of the virtual smart card when the key is created.
    After creating the key, start the agent with Start-SmartCardSSHAgent.ps1.
.PARAMETER ReaderName
    Name of the smart card reader of the virtual smart card,
    as printed by New-TPMVirtualSmartCard.ps1 and Get-TPMVirtualSmartCard.ps1
.PARAMETER KeyName
    Name of the key container and the subject of the certificate.
    It may contain only letters, digits and the characters ".", "_", "@" and "-",
    and it may be at most 39 characters long, since that is the maximum length of a smart card container name.
.PARAMETER Thumbprint
    Use an existing certificate with this thumbprint instead of creating a new key.
    The certificate must be in the personal certificate store of the current user,
    and its key must be on a smart card.
.PARAMETER Comment
    Comment of the OpenSSH public key
.PARAMETER PublicKeyPath
    Path for the OpenSSH public key
.PARAMETER Force
    Overwrite an existing public key file
.EXAMPLE
    .\New-SmartCardSSHKey.ps1 -ReaderName "Microsoft Virtual Smart Card 0"
.EXAMPLE
    .\New-SmartCardSSHKey.ps1 -Thumbprint "3F7A9C2E81D45B06E9C3A7F1D28B54E0C6A91F3D"
#>

[CmdletBinding(DefaultParameterSetName="Create")]
param(
    [Parameter(ParameterSetName="Create")][string]$ReaderName,
    [Parameter(ParameterSetName="Create")][string]$KeyName = "${env:USERNAME}@$(${env:COMPUTERNAME}.ToLower())_tpm-vsc",
    [Parameter(ParameterSetName="Existing", Mandatory=$true)][ValidatePattern('^[0-9A-Fa-f]{40}$')][string]$Thumbprint,
    [ValidatePattern('^\S+$')][string]$Comment = "${env:USERNAME}@$(${env:COMPUTERNAME}.ToLower())_tpm-vsc",
    [string]$PublicKeyPath = "${HOME}\.ssh\id_rsa_tpm_vsc_$(${env:COMPUTERNAME}.ToLower()).pub",
    [switch]$Force
)

Set-StrictMode -Version 3.0
. "${PSScriptRoot}\..\Utils.ps1"
. "${PSScriptRoot}\SmartCardUtils.ps1"

if (Test-Admin) {
    Show-Output "Run this script as a normal user, not elevated." -ForegroundColor Red
    exit 1
}
if (-not (Test-SmartCardSSHRequirement)) {
    exit 1
}
if ((Test-Path "${PublicKeyPath}") -and (-not $Force)) {
    Show-Output "The public key file `"${PublicKeyPath}`" already exists. Use -Force to overwrite it." -ForegroundColor Red
    exit 1
}
# The default values are not checked by the validation attributes, so check them here.
if ($Comment -notmatch '^\S+$') {
    Show-Output "The comment `"${Comment}`" must not contain whitespace." -ForegroundColor Red
    exit 1
}

if ($Thumbprint) {
    $Cert = Get-Item -Path "Cert:\CurrentUser\My\$($Thumbprint.ToUpper())" -ErrorAction SilentlyContinue
    if ($null -eq $Cert) {
        Show-Output "The certificate `"${Thumbprint}`" was not found in the personal certificate store." -ForegroundColor Red
        exit 1
    }
    $KeyName = $Cert.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName, $false)
} else {
    $Readers = @(Get-SmartCardReader)
    if (-not $ReaderName) {
        Show-Output "Smart card readers:"
        for ($i = 0; $i -lt $Readers.Count; $i++) {
            Show-Output "  [${i}] $($Readers[$i])"
        }
        $Index = Read-Host "Select the reader of the TPM virtual smart card"
        if ($Index -notmatch "^\d+$" -or [int]$Index -ge $Readers.Count) {
            Show-Output "Invalid selection." -ForegroundColor Red
            exit 1
        }
        $ReaderName = $Readers[[int]$Index]
    }
    if ($KeyName -notmatch '^[A-Za-z0-9._@-]{1,39}$') {
        Show-Output "The key name `"${KeyName}`" may contain only letters, digits and the characters `".`", `"_`", `"@`" and `"-`", and it may be at most 39 characters long. Use -KeyName to give another name." -ForegroundColor Red
        exit 1
    }
    if ($Readers -notcontains $ReaderName) {
        Show-Output "The reader `"${ReaderName}`" was not found." -ForegroundColor Red
        exit 1
    }

    Show-Output "Creating the key `"${KeyName}`" on the card in `"${ReaderName}`". Enter the PIN of the card when asked." -ForegroundColor Cyan
    $CertParams = @{
        Subject = "CN=${KeyName}"
        Type = "Custom"
        Provider = $SmartCardKSP
        # The smart card KSP accepts the reader name as a prefix of the container name.
        Container = "\\.\${ReaderName}\${KeyName}"
        KeyAlgorithm = "RSA"
        KeyLength = 2048
        KeyUsage = "DigitalSignature"
        HashAlgorithm = "SHA256"
        NotAfter = (Get-Date).AddYears(10)
        CertStoreLocation = "Cert:\CurrentUser\My"
    }
    $Cert = New-SelfSignedCertificate @CertParams
    if ($null -eq $Cert) {
        Show-Output "Creating the key failed." -ForegroundColor Red
        exit 1
    }
    Show-Output "Created the certificate with the thumbprint $($Cert.Thumbprint)."
}

# Ensure that the key is on a smart card, and not e.g. in a software key store.
$Key = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Cert)
if (-not ($Key -is [System.Security.Cryptography.RSACng]) -or $Key.Key.Provider.Provider -ne $SmartCardKSP) {
    Show-Output "The private key of the certificate is not on a smart card." -ForegroundColor Red
    exit 1
}

$PublicKey = ConvertTo-OpenSSHPublicKey -Certificate $Cert
Set-Content -Path "${PublicKeyPath}" -Value "${PublicKey} ${Comment}" -Encoding ASCII
New-Item -ItemType Directory -Path "${SmartCardSSHDir}" -Force | Out-Null
@{
    KeyName = $KeyName
    Thumbprint = $Cert.Thumbprint
    PublicKeyPath = $PublicKeyPath
} | ConvertTo-Json | Set-Content -Path "${SmartCardSSHConfig}" -Encoding UTF8

$PublicKeyName = Split-Path -Leaf $PublicKeyPath
$SocketName = Split-Path -Leaf $SmartCardSSHSocket
Show-Output "The public key was saved to `"${PublicKeyPath}`":" -ForegroundColor Green
Show-Output "${PublicKey} ${Comment}"
Show-Output "Add the public key to GitHub, and add the following to the SSH configuration before the other identities:"
Show-Output @"
Host github.com
    IdentityAgent ~/.ssh/agent/${SocketName}
    IdentityFile ~/.ssh/${PublicKeyName}
    IdentitiesOnly yes
"@
Show-Output "Then start the agent with .\Start-SmartCardSSHAgent.ps1"
