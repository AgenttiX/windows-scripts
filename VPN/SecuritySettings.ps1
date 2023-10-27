<#
.SYNOPSIS
    Security settings for IKEv2 VPN
.DESCRIPTION
    The VPN security settings are needed in several scripts, and therefore it's the easiest to store them in one place.
.LINK
    https://directaccess.richardhicks.com/2018/12/10/always-on-vpn-ikev2-security-configuration/
.LINK
    https://docs.microsoft.com/en-us/azure/vpn-gateway/vpn-gateway-about-compliance-crypto
.LINK
    https://learn.microsoft.com/en-us/powershell/module/vpnclient/set-vpnconnectionipsecconfiguration
#>

Write-Host "Loading VPN security settings from file."

# "Specifies authentication header (AH) transform in the IPsec policy."
$AuthenticationTransformConstants = "SHA256128"
# $AuthenticationTransformConstants = "GCMAES256"
# "Specifies Encapsulating Security Payload (ESP) cipher transform in the IPsec policy."
$CipherTransformConstants = "AES256"
# $CipherTransformConstants = "GCMAES256"
# By default Windows uses Group2, which is 1024-bit and therefore not secure.
# https://weakdh.org/
# As of 2023, Group14 (2048-bit MODP) is the bare minimum acceptable.
# $DHGroup = "Group14"
# The Group24 (2048-bit MODP with a 256-bit prime order subgroup) provides a security level of 112 bits
# for symmetric encryption, and is therefore not sufficient for AES-256.
# https://datatracker.ietf.org/doc/html/rfc5114#section-4
# As of 2023, ECP384 is the strongest supported by Windows 11.
$DHGroup = "ECP384"
# Specifying GCMAES256 causes VPN_Profile.ps1 to fail with the error message
# "A general error occurred that is not covered by a more specific error code."
# This seems to be a bug in Windows 11.
# GCMAES cannot be combined with AES, and therefore the other settings have to use AES-CCM as well.
# $EncryptionMethod = "GCMAES256"
$EncryptionMethod = "AES256"
# $IntegrityCheckMethod = "SHA256"
$IntegrityCheckMethod = "SHA384"
# "DHGroup2048 & PFS2048 are the same as Diffie-Hellman Group 14 in IKE and IPsec PFS."
# $PfsGroup = "PFS2048"
$PfsGroup = "ECP384"

# "IKEv2 Main Mode SA lifetime is fixed at 28,800 seconds on the Azure VPN gateways."
$SALifeTimeSeconds = 28800
$MMSALifeTimeSeconds = 86400
$SADataSizeForRenegotiationKilobytes = 1024000
