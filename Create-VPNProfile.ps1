<#
.SYNOPSIS
    Create a connection profile for the Vexlum VPN using a computer that already has the VPN configured.
.NOTES
    This script cannot be run over Remote Desktop or in a Hyper-V enhanced session.
.LINK
    https://docs.microsoft.com/en-us/windows-server/remote/remote-access/vpn/always-on-vpn/deploy/vpn-deploy-client-vpn-connections
#>

param(
    [Parameter(Mandatory=$true)][string]$TemplateName,
    [Parameter(Mandatory=$true)][string]$ProfileName,
    # VPN/RAS server, not NPS
    [Parameter(Mandatory=$true)][string]$Servers,
    [Parameter(Mandatory=$true)][string]$DNSSuffix,
    [Parameter(Mandatory=$true)][string]$DomainName,
    [Parameter(Mandatory=$true)][string]$DNSServers,
    [Parameter(Mandatory=$true)][string]$TrustedNetwork
)

# The $env:USERPROFILE does not work if OneDrive sync is enabled for the user folder
# https://stackoverflow.com/a/64256803
$DesktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
# Write-Host $DesktopPath

$Connection = Get-VpnConnection -Name $TemplateName
if(!$Connection)
{
    $Message = "Unable to get $TemplateName connection profile: $_"
    Write-Host "$Message"
    exit
}
$EAPSettings= $Connection.EapConfigXmlStream.InnerXml

$ProfileXML = @("
<VPNProfile>
  <DnsSuffix>$DNSSuffix</DnsSuffix>
  <NativeProfile>
<Servers>$Servers</Servers>
<NativeProtocolType>IKEv2</NativeProtocolType>
<Authentication>
  <UserMethod>Eap</UserMethod>
  <Eap>
   <Configuration>
    $EAPSettings
   </Configuration>
  </Eap>
</Authentication>
<RoutingPolicyType>SplitTunnel</RoutingPolicyType>
  </NativeProfile>
<AlwaysOn>true</AlwaysOn>
<RememberCredentials>true</RememberCredentials>
<TrustedNetworkDetection>$TrustedNetwork</TrustedNetworkDetection>
  <DomainNameInformation>
<DomainName>$DomainName</DomainName>
<DnsServers>$DNSServers</DnsServers>
</DomainNameInformation>
</VPNProfile>
")

$ProfileXML | Out-File -FilePath "${DesktopPath}\VPN_Profile.xml"
$Script = @("
`$ProfileName = '$ProfileName'
`$ProfileNameEscaped = `$ProfileName -replace ' ', '%20'

`$ProfileXML = '$ProfileXML'

`$ProfileXML = `$ProfileXML -replace '<', '&lt;'
`$ProfileXML = `$ProfileXML -replace '>', '&gt;'
`$ProfileXML = `$ProfileXML -replace '`"', '&quot;'

`$nodeCSPURI = `"./Vendor/MSFT/VPNv2`"
`$namespaceName = `"root\cimv2\mdm\dmmap`"
`$className = `"MDM_VPNv2_01`"

try
{
    `$username = Gwmi -Class Win32_ComputerSystem | select username
    `$objuser = New-Object System.Security.Principal.NTAccount(`$username.username)
    `$sid = `$objuser.Translate([System.Security.Principal.SecurityIdentifier])
    `$SidValue = `$sid.Value
    `$Message = `"User SID is `$SidValue.`"
    Write-Host `"`$Message`"
}
catch [Exception]
{
    `$Message = `"Unable to get user SID. User may be logged on over Remote Desktop: `$_`"
    Write-Host `"`$Message`"
    exit
}

`$session = New-CimSession
`$options = New-Object Microsoft.Management.Infrastructure.Options.CimOperationOptions
`$options.SetCustomOption(`"PolicyPlatformContext_PrincipalContext_Type`", `"PolicyPlatform_UserContext`", `$false)
`$options.SetCustomOption(`"PolicyPlatformContext_PrincipalContext_Id`", `"`$SidValue`", `$false)

try
{
    `$deleteInstances = `$session.EnumerateInstances(`$namespaceName, `$className, `$options)
    foreach (`$deleteInstance in `$deleteInstances)
    {
        `$InstanceId = `$deleteInstance.InstanceID
        if (`"`$InstanceId`" -eq `"`$ProfileNameEscaped`")
        {
            `$session.DeleteInstance(`$namespaceName, `$deleteInstance, `$options)
            `$Message = `"Removed `$ProfileName profile `$InstanceId`"
            Write-Host `"`$Message`"
        } else {
            `$Message = `"Ignoring existing VPN profile `$InstanceId`"
            Write-Host `"`$Message`"
        }
    }
}
catch [Exception]
{
    `$Message = `"Unable to remove existing outdated instance(s) of `$ProfileName profile: `$_`"
    Write-Host `"`$Message`"
    exit
}

try
{
    `$newInstance = New-Object Microsoft.Management.Infrastructure.CimInstance `$className, `$namespaceName
    `$property = [Microsoft.Management.Infrastructure.CimProperty]::Create(`"ParentID`", `"`$nodeCSPURI`", `"String`", `"Key`")
    `$newInstance.CimInstanceProperties.Add(`$property)
    `$property = [Microsoft.Management.Infrastructure.CimProperty]::Create(`"InstanceID`", `"`$ProfileNameEscaped`", `"String`",      `"Key`")
    `$newInstance.CimInstanceProperties.Add(`$property)
    `$property = [Microsoft.Management.Infrastructure.CimProperty]::Create(`"ProfileXML`", `"`$ProfileXML`", `"String`", `"Property`")
    `$newInstance.CimInstanceProperties.Add(`$property)
    `$session.CreateInstance(`$namespaceName, `$newInstance, `$options)
    `$Message = `"Created `$ProfileName profile.`"

Write-Host `"`$Message`"
}
catch [Exception]
{
    `$Message = `"Unable to create `$ProfileName profile: `$_`"
    Write-Host `"`$Message`"
    exit
}

`$Message = `"Script Complete`"
Write-Host `"`$Message`"
")

$ScriptPath = "${DesktopPath}\VPN_Profile.ps1"
$Script | Out-File -FilePath "${ScriptPath}"
. ".\Sign-Script.ps1" -FilePath "${ScriptPath}"

$Message = "Successfully created VPN_Profile.xml and VPN_Profile.ps1 on the desktop."
Write-Host "$Message"
