<#
.SYNOPSIS
    Fix the security parameters of a Windows IKEv2 VPN connection.
.DESCRIPTION
    Windows uses insecure Diffie-Hellmann parameters and other security settings by default. This script fixes it.
.LINK
    https://docs.microsoft.com/fi-fi/windows/security/identity-protection/vpn/how-to-configure-diffie-hellman-protocol-over-ikev2-vpn-connections
.LINK
    https://directaccess.richardhicks.com/2018/12/10/always-on-vpn-ikev2-security-configuration/
.LINK
    https://weakdh.org/
.LINK
    https://docs.strongswan.org/docs/5.9/interop/windowsClients.html#strong_ke
#>
param(
    [string]$ConnectionName = $null,
    [switch]$Reset = $false
)
. "${PSScriptRoot}\SecuritySettings.ps1"

$ProductType = (Get-CimInstance -ClassName Win32_OperatingSystem).ProductType

if (($ProductType -eq 2) -or ($ProductType -eq 3)) {
    Write-Output "This seems to be a server. Configuring VPN server."
    if ($Reset) {
        Write-Output "Resetting to default settings."
        Set-VpnServerConfiguration -RevertToDefault -PassThru
        Restart-Service RemoteAccess -PassThru
        return
    }
    Set-VpnServerConfiguration `
        -TunnelType IKEv2 `
        -CustomPolicy `
        -AuthenticationTransformConstants "${AuthenticationTransformConstants}" `
        -CipherTransformConstants "${CipherTransformConstants}" `
        -DHGroup "${DHGroup}" `
        -EncryptionMethod "${EncryptionMethod}" `
        -IntegrityCheckMethod "${IntegrityCheckMethod}" `
        -PfsGroup "${PfsGroup}" `
        -SALifeTimeSeconds "${SALifeTimeSeconds}" `
        -MMSALifeTimeSeconds "${MMSALifeTimeSeconds}" `
        -SADataSizeForRenegotiationKilobytes "${SADataSizeForRenegotiationKilobytes}" `
        -PassThru
    Restart-Service RemoteAccess -PassThru
} elseif ($ProductType -eq 1) {
    Write-Output "This seems to be a workstation. Configuring VPN client."
    if ($null -eq $ConnectionName) {
        Write-Output "Cannot configure a client without connection name. Please provide the ConnectionName argument."
        return
    }
    if ($Reset) {
        Write-Output "Resetting to default settings."
        Set-VpnConnectionIPsecConfiguration -ConnectionName "${ConnectionName}" -RevertToDefault -PassThru -Force
    }
    Set-VpnConnectionIPsecConfiguration `
        -ConnectionName "${ConnectionName}" `
        -AuthenticationTransformConstants "${AuthenticationTransformConstants}" `
        -CipherTransformConstants "${CipherTransformConstants}" `
        -DHGroup "${DHGroup}" `
        -EncryptionMethod "${EncryptionMethod}" `
        -IntegrityCheckMethod "${IntegrityCheckMethod}" `
        -PfsGroup "${PfsGroup}" `
        -PassThru `
        -Force
} else {
    Write-Output "Got unknown product type: ${ProductType}. Aborting."
}
Write-Output "VPN configuration ready."
