<#
.SYNOPSIS
    Create an SSH key on a TPM virtual smart card
.DESCRIPTION
    Creates a non-exportable RSA key and a self-signed certificate for it on the virtual smart card
    created by New-TPMVirtualSmartCard.ps1, and saves the public key in the OpenSSH format.
    The key is accessed from SSH using the OpenSC PKCS#11 module, so OpenSC has to be installed first.
    It can be installed with Install-Software.ps1, or with either winget or Chocolatey:
    winget install --exact --id OpenSC.OpenSC
    choco install opensc

    Run this as a normal user, since the certificate is stored in the certificate store of the current user.
    Windows asks for the PIN of the virtual smart card when the key is created.
    After creating the key, start the agent with Start-SmartCardSSHAgent.ps1.
.PARAMETER ReaderName
    Name of the smart card reader of the virtual smart card,
    as printed by New-TPMVirtualSmartCard.ps1 and Get-TPMVirtualSmartCard.ps1
.PARAMETER KeyName
    Name of the key container and the subject of the certificate
.PARAMETER PublicKeyPath
    Path for the OpenSSH public key
.PARAMETER Force
    Overwrite an existing public key file
#>

param(
    [string]$ReaderName,
    [string]$KeyName = "ssh-$(${env:COMPUTERNAME}.ToLower())",
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
if (-not (Test-OpenSC)) {
    exit 1
}
if ((Test-Path "${PublicKeyPath}") -and (-not $Force)) {
    Show-Output "The public key file `"${PublicKeyPath}`" already exists. Use -Force to overwrite it." -ForegroundColor Red
    exit 1
}

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

# OpenSC reads the public key from the certificate on the card, so ensure that the certificate is stored there.
$Key = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Cert)
if (-not ($Key -is [System.Security.Cryptography.RSACng])) {
    Show-Output "The private key is not a CNG key." -ForegroundColor Red
    exit 1
}
$CertProperty = "SmartCardKeyCertificate"
if (-not $Key.Key.HasProperty($CertProperty, [System.Security.Cryptography.CngPropertyOptions]::None)) {
    Show-Output "Writing the certificate to the card."
    $Key.Key.SetProperty((New-Object System.Security.Cryptography.CngProperty($CertProperty, $Cert.RawData, [System.Security.Cryptography.CngPropertyOptions]::None)))
}

Set-OpenSCConfig -ReaderName $ReaderName
if (-not (Test-SingleToken)) {
    Show-Output "Edit the ignored_readers in `"${SmartCardSSHOpenSCConf}`" and run the script again with -Force." -ForegroundColor Red
    exit 1
}

$PublicKey = ConvertTo-OpenSSHPublicKey -Certificate $Cert
$OldConf = $env:OPENSC_CONF
$env:OPENSC_CONF = $SmartCardSSHOpenSCConf
try {
    # This only reads the public keys and does not log in to the card.
    $PKCS11Keys = @(& "${GitUsrBin}\ssh-keygen.exe" -D (ConvertTo-MsysPath $OpenSCModule) 2>$null)
} finally {
    $env:OPENSC_CONF = $OldConf
}
if (-not ($PKCS11Keys | Where-Object { $_.StartsWith("${PublicKey} ") -or $_ -eq $PublicKey })) {
    Show-Output "The key was not found using OpenSC. The keys found were:" -ForegroundColor Red
    $PKCS11Keys | ForEach-Object { Show-Output "  $_" }
    exit 1
}

Set-Content -Path "${PublicKeyPath}" -Value "${PublicKey} ${KeyName}" -Encoding ASCII
@{
    ReaderName = $ReaderName
    KeyName = $KeyName
    Thumbprint = $Cert.Thumbprint
    PublicKeyPath = $PublicKeyPath
} | ConvertTo-Json | Set-Content -Path "${SmartCardSSHConfig}" -Encoding UTF8

$PublicKeyName = Split-Path -Leaf $PublicKeyPath
$SocketName = Split-Path -Leaf $SmartCardSSHSocket
Show-Output "The public key was saved to `"${PublicKeyPath}`":" -ForegroundColor Green
Show-Output "${PublicKey} ${KeyName}"
Show-Output "Add the public key to GitHub, and add the following to the SSH configuration before the other identities:"
Show-Output @"
Host github.com
    IdentityAgent ~/.ssh/agent/${SocketName}
    IdentityFile ~/.ssh/${PublicKeyName}
    IdentitiesOnly yes
"@
Show-Output "Then start the agent with .\Start-SmartCardSSHAgent.ps1"
