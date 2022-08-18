<#
.SYNOPSIS
    Security settings for IKEv2 VPN
.DESCRIPTION
    The VPN security settings are needed in several scripts, and therefore it's the easiest to store them in one place.
.LINK
    https://directaccess.richardhicks.com/2018/12/10/always-on-vpn-ikev2-security-configuration/
.LINK
    https://docs.microsoft.com/en-us/azure/vpn-gateway/vpn-gateway-about-compliance-crypto
#>

Write-Host "Loading VPN security settings from file."

# "Specifies authentication header (AH) transform in the IPsec policy."
$AuthenticationTransformConstants = "SHA256128"
# "Specifies Encapsulating Security Payload (ESP) cipher transform in the IPsec policy."
$CipherTransformConstants = "AES128"
# By default Windows uses Group2, which is 1024-bit and therefore not secure.
# https://weakdh.org/
$DHGroup = "Group14"
$EncryptionMethod = "AES128"
$IntegrityCheckMethod = "SHA256"
# "DHGroup2048 & PFS2048 are the same as Diffie-Hellman Group 14 in IKE and IPsec PFS."
$PfsGroup = "PFS2048"
# "IKEv2 Main Mode SA lifetime is fixed at 28,800 seconds on the Azure VPN gateways."
$SALifeTimeSeconds = 28800
$MMSALifeTimeSeconds = 86400
$SADataSizeForRenegotiationKilobytes = 1024000
