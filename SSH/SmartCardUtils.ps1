<#
.SYNOPSIS
    Shared helpers for the TPM virtual smart card SSH scripts
.DESCRIPTION
    The TPM virtual smart card is accessed from SSH using the OpenSC PKCS#11 module.
    OpenSSH sends the PIN given to "ssh-add -s" to every token that the PKCS#11 module exposes.
    Therefore OpenSC is used with a dedicated configuration file,
    which only allows the GIDS card driver and ignores the other readers,
    so that e.g. the PIV applet of a YubiKey does not receive the PIN of the virtual smart card.
.LINK
    https://github.com/openssh/openssh-portable/blob/master/ssh-pkcs11.c
.LINK
    https://github.com/OpenSC/OpenSC/blob/master/etc/opensc.conf.example.in
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute(
    "PSUseDeclaredVarsMoreThanAssignments",
    "SmartCardKSP",
    Justification="Used in scripts"
)]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute(
    "PSUseDeclaredVarsMoreThanAssignments",
    "SmartCardSSHSocket",
    Justification="Used in scripts"
)]
param()

$GitUsrBin = "${env:ProgramFiles}\Git\usr\bin"
$OpenSCDir = "${env:ProgramFiles}\OpenSC Project\OpenSC"
$OpenSCModule = "${OpenSCDir}\pkcs11\opensc-pkcs11.dll"
$OpenSCTool = "${OpenSCDir}\tools\pkcs11-tool.exe"
$SmartCardKSP = "Microsoft Smart Card Key Storage Provider"
$SmartCardSSHDir = "${env:LOCALAPPDATA}\SmartCardSSH"
$SmartCardSSHConfig = "${SmartCardSSHDir}\config.json"
$SmartCardSSHOpenSCConf = "${SmartCardSSHDir}\opensc.conf"
# The ~/.ssh/agent directory is also used by OpenSSH for the default agent sockets.
$SmartCardSSHSocket = "${HOME}\.ssh\agent\tpm-vsc.sock"

function Get-SmartCardReader {
    <#
    .SYNOPSIS
        List the PC/SC smart card reader names
    .LINK
        https://learn.microsoft.com/en-us/windows/win32/api/winscard/nf-winscard-scardlistreadersw
    #>
    [OutputType([string[]])]
    param()

    if (-not ("SmartCardSSH.WinSCard" -as [type])) {
        Add-Type -Namespace "SmartCardSSH" -Name "WinSCard" -MemberDefinition @"
[DllImport("winscard.dll")]
public static extern int SCardEstablishContext(uint dwScope, IntPtr pvReserved1, IntPtr pvReserved2, out IntPtr phContext);
[DllImport("winscard.dll")]
public static extern int SCardReleaseContext(IntPtr hContext);
[DllImport("winscard.dll", CharSet = CharSet.Unicode)]
public static extern int SCardListReadersW(IntPtr hContext, string mszGroups, char[] mszReaders, ref uint pcchReaders);
"@
    }
    # SCARD_SCOPE_USER
    $Context = [IntPtr]::Zero
    $Result = [SmartCardSSH.WinSCard]::SCardEstablishContext(0, [IntPtr]::Zero, [IntPtr]::Zero, [ref]$Context)
    if ($Result -ne 0) {
        throw ("SCardEstablishContext failed with 0x{0:X8}" -f $Result)
    }
    try {
        [uint32]$Length = 0
        $Result = [SmartCardSSH.WinSCard]::SCardListReadersW($Context, $null, $null, [ref]$Length)
        # SCARD_E_NO_READERS_AVAILABLE
        if ($Result -eq 0x8010002E) {
            return @()
        }
        if ($Result -ne 0) {
            throw ("SCardListReaders failed with 0x{0:X8}" -f $Result)
        }
        $Buffer = New-Object char[] $Length
        $Result = [SmartCardSSH.WinSCard]::SCardListReadersW($Context, $null, $Buffer, [ref]$Length)
        if ($Result -ne 0) {
            throw ("SCardListReaders failed with 0x{0:X8}" -f $Result)
        }
        return @((New-Object string (, $Buffer)).Split([char]0) | Where-Object { $_ -ne "" })
    } finally {
        [void][SmartCardSSH.WinSCard]::SCardReleaseContext($Context)
    }
}

function Get-ElevateCommand {
    <#
    .SYNOPSIS
        Get the command for re-running a script elevated with the same arguments
    .DESCRIPTION
        Elevate in Utils.ps1 does not pass the arguments of the script on to the elevated process.
        The values are single-quoted, so they must not contain quotes.
    #>
    [OutputType([string])]
    param(
        [Parameter(Mandatory=$true)][string]$ScriptPath,
        [Parameter(Mandatory=$true)][AllowEmptyCollection()][System.Collections.IDictionary]$BoundParameters
    )
    $Arguments = @()
    foreach ($Parameter in $BoundParameters.GetEnumerator()) {
        if ($Parameter.Key -eq "Elevated") {
            continue
        }
        if ($Parameter.Value -is [System.Management.Automation.SwitchParameter]) {
            if ($Parameter.Value.IsPresent) {
                $Arguments += "-$($Parameter.Key)"
            }
        } else {
            $Arguments += "-$($Parameter.Key) '$($Parameter.Value)'"
        }
    }
    return ("& '{0}' {1}" -f ($ScriptPath, ($Arguments -join " "))).TrimEnd()
}

function Get-TPMVirtualSmartCardDevice {
    <#
    .SYNOPSIS
        List the TPM virtual smart cards
    .DESCRIPTION
        The name of the card is the friendly name of its reader device,
        but the PC/SC reader name is of the format "Microsoft Virtual Smart Card <n>".
        The PC/SC reader name is determined from the child devices of the reader device.
    #>
    [OutputType([PSCustomObject[]])]
    param()
    $Devices = @(Get-PnpDevice -InstanceId "ROOT\SMARTCARDREADER\*" -ErrorAction SilentlyContinue)
    foreach ($Device in $Devices) {
        $ReaderName = $null
        $Children = Get-PnpDeviceProperty -InstanceId $Device.InstanceId -KeyName "DEVPKEY_Device_Children" -ErrorAction SilentlyContinue
        if ($null -ne $Children -and $null -ne $Children.Data) {
            foreach ($Child in $Children.Data) {
                if ($Child -match "&(Microsoft_Virtual_Smart_Card_\d+)_SCFILTER") {
                    $ReaderName = $Matches[1].Replace("_", " ")
                    break
                }
            }
        }
        [PSCustomObject]@{
            Name = $Device.FriendlyName
            InstanceId = $Device.InstanceId
            ReaderName = $ReaderName
            Status = $Device.Status
        }
    }
}

function Get-SmartCardSSHConfig {
    <#
    .SYNOPSIS
        Load the configuration saved by New-SmartCardSSHKey.ps1
    #>
    [OutputType([PSCustomObject])]
    param()
    if (-not (Test-Path "${SmartCardSSHConfig}")) {
        throw "The configuration was not found at `"${SmartCardSSHConfig}`". Run New-SmartCardSSHKey.ps1 first."
    }
    return Get-Content -Path "${SmartCardSSHConfig}" -Raw | ConvertFrom-Json
}

function ConvertTo-MsysPath {
    <#
    .SYNOPSIS
        Convert a Windows path to the format used by the OpenSSH of Git for Windows
    #>
    [OutputType([string])]
    param(
        [Parameter(Mandatory=$true)][string]$Path
    )
    return (& "${GitUsrBin}\cygpath.exe" -u "${Path}")
}

function ConvertTo-OpenSSHPublicKey {
    <#
    .SYNOPSIS
        Convert the RSA public key of a certificate to the OpenSSH public key format
    .LINK
        https://datatracker.ietf.org/doc/html/rfc4253#section-6.6
    #>
    [OutputType([string])]
    param(
        [Parameter(Mandatory=$true)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )
    $RSA = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($Certificate)
    if ($null -eq $RSA) {
        throw "The certificate does not have an RSA public key."
    }
    $Parameters = $RSA.ExportParameters($false)

    $Stream = New-Object System.IO.MemoryStream
    $WriteBlock = {
        param([byte[]]$Data)
        $LengthBytes = [BitConverter]::GetBytes([uint32]$Data.Length)
        [Array]::Reverse($LengthBytes)
        $Stream.Write($LengthBytes, 0, 4)
        $Stream.Write($Data, 0, $Data.Length)
    }
    $WriteMpint = {
        param([byte[]]$Data)
        # An mpint with the most significant bit set must be prefixed with a zero byte.
        if ($Data[0] -band 0x80) {
            $Data = [byte[]](, 0 + $Data)
        }
        & $WriteBlock $Data
    }
    & $WriteBlock ([System.Text.Encoding]::ASCII.GetBytes("ssh-rsa"))
    & $WriteMpint $Parameters.Exponent
    & $WriteMpint $Parameters.Modulus
    return "ssh-rsa " + [Convert]::ToBase64String($Stream.ToArray())
}

function Set-OpenSCConfig {
    <#
    .SYNOPSIS
        Write an OpenSC configuration that exposes only the TPM virtual smart card
    #>
    param(
        [Parameter(Mandatory=$true)][string]$ReaderName
    )
    # Ignore all other readers that are currently present, and also the ones that may appear later.
    # The matching is case-sensitive and by substring, so the reader index is stripped.
    $Ignored = @("Yubico", "YubiKey", "Windows Hello")
    foreach ($Reader in Get-SmartCardReader) {
        if ($Reader -ne $ReaderName) {
            $Ignored += ($Reader -replace "\s+\d+$", "")
        }
    }
    # A substring of the virtual smart card reader name must not be ignored.
    $Ignored = @($Ignored | Sort-Object -Unique | Where-Object { -not $ReaderName.Contains($_) })
    $IgnoredList = ($Ignored | ForEach-Object { "`"" + ($_ -replace '"', '') + "`"" }) -join ", "

    New-Item -ItemType Directory -Path "${SmartCardSSHDir}" -Force | Out-Null
    $Conf = @"
# Generated by SmartCardUtils.ps1 of windows-scripts.
# This configuration is used only for the SSH agent of the TPM virtual smart card.
app default {
    # Microsoft TPM virtual smart cards use the GIDS profile.
    card_drivers = gids;
    ignored_readers = ${IgnoredList};
}
"@
    Set-Content -Path "${SmartCardSSHOpenSCConf}" -Value $Conf -Encoding ASCII
}

function Test-SingleToken {
    <#
    .SYNOPSIS
        Check that the OpenSC PKCS#11 module exposes exactly one token
    .DESCRIPTION
        This must be checked before giving the PIN to ssh-add,
        since OpenSSH attempts to log in to every token with the same PIN.
        pkcs11-tool --list-token-slots does not log in to the tokens.
    #>
    [OutputType([bool])]
    param()
    $OldConf = $env:OPENSC_CONF
    $env:OPENSC_CONF = $SmartCardSSHOpenSCConf
    try {
        $Output = & "${OpenSCTool}" --module "${OpenSCModule}" --list-token-slots 2>&1 | Out-String
    } finally {
        $env:OPENSC_CONF = $OldConf
    }
    $TokenCount = ([regex]::Matches($Output, "token label")).Count
    if ($TokenCount -ne 1) {
        Show-Information "OpenSC exposes ${TokenCount} tokens, but exactly one is required:" -ForegroundColor Red
        Show-Information $Output
        return $false
    }
    return $true
}

function Test-OpenSC {
    <#
    .SYNOPSIS
        Check that OpenSC and Git for Windows are installed
    #>
    [OutputType([bool])]
    param()
    if (-not (Test-Path "${GitUsrBin}\ssh-agent.exe")) {
        Show-Information "Git for Windows was not found at `"${GitUsrBin}`"." -ForegroundColor Red
        return $false
    }
    if (-not ((Test-Path "${OpenSCModule}") -and (Test-Path "${OpenSCTool}"))) {
        Show-Information "OpenSC was not found at `"${OpenSCDir}`". Install it with Install-Software.ps1, or with either of:" -ForegroundColor Red
        Show-Information "winget install --exact --id OpenSC.OpenSC"
        Show-Information "choco install opensc"
        return $false
    }
    return $true
}
