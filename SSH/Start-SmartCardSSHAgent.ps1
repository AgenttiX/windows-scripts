<#
.SYNOPSIS
    Start the SSH agent for the key on the TPM virtual smart card
.DESCRIPTION
    Starts the Pageant of PuTTY-CAC with the certificate of the key on the TPM virtual smart card,
    and ssh-pageant of Git for Windows with a fixed socket path,
    so that the OpenSSH of Git for Windows can use Pageant through IdentityAgent in the SSH configuration.
    This works also for programs that do not have SSH_AUTH_SOCK set, such as IDEs.

    Pageant is started with PIN caching enabled and signing confirmation prompts disabled.
    PuTTY-CAC saves these settings to the registry, so they also apply when Pageant is started otherwise.
    The script makes a test signature, so that the PIN is asked immediately,
    and not later when a background process such as an IDE uses the key.
    After that the key can be used without prompts until Pageant is stopped or the user logs out.

    Note that while the key is loaded, any process of the current user can use it through Pageant.
    The key cannot be copied from the TPM, however.
.PARAMETER Restart
    Stop Pageant and ssh-pageant before starting them again. This removes all keys from Pageant.
.PARAMETER Stop
    Only stop Pageant and ssh-pageant. This removes all keys from Pageant.
#>

param(
    [switch]$Restart,
    [switch]$Stop
)

Set-StrictMode -Version 3.0
. "${PSScriptRoot}\..\Utils.ps1"
. "${PSScriptRoot}\SmartCardUtils.ps1"

$BridgePidPath = "${SmartCardSSHDir}\ssh-pageant.pid"

function Stop-Bridge {
    <#
    .SYNOPSIS
        Stop ssh-pageant and remove its socket
    #>
    if (Test-Path "${BridgePidPath}") {
        # ssh-pageant -k uses the MSYS process ID, which differs from the Windows process ID.
        $env:SSH_PAGEANT_PID = (Get-Content -Path "${BridgePidPath}" -Raw).Trim()
        & "${GitUsrBin}\ssh-pageant.exe" -k *> $null
        Remove-Item -Path "${BridgePidPath}"
        Remove-Item Env:\SSH_PAGEANT_PID
    }
    # MSYS2 creates the socket as a file with the system attribute, so -Force is required.
    if (Test-Path "${SmartCardSSHSocket}") {
        Remove-Item -Path "${SmartCardSSHSocket}" -Force
    }
}

function Test-KeyLoaded {
    <#
    .SYNOPSIS
        Check whether the key is available through the socket
    #>
    [OutputType([bool])]
    param()
    $LoadedKeys = @(& "${GitUsrBin}\ssh-add.exe" -L 2>$null)
    return [bool]($LoadedKeys | Where-Object { $_.StartsWith($PublicKey) })
}

if (Test-Admin) {
    Show-Output "Run this script as a normal user, not elevated." -ForegroundColor Red
    exit 1
}
if (-not (Test-SmartCardSSHRequirement)) {
    exit 1
}
$Config = Get-SmartCardSSHConfig
$PublicKey = ((Get-Content -Path $Config.PublicKeyPath -Raw).Trim() -split " ")[0..1] -join " "

$OldAuthSock = $env:SSH_AUTH_SOCK
$env:SSH_AUTH_SOCK = ConvertTo-MsysPath $SmartCardSSHSocket
try {
    if ($Stop -or $Restart) {
        Show-Output "Stopping Pageant and ssh-pageant."
        Stop-Bridge
        Get-Process -Name "pageant" -ErrorAction SilentlyContinue | Stop-Process
        if ($Stop) {
            exit 0
        }
        Start-Sleep -Seconds 1
    }

    if (-not (Get-Process -Name "pageant" -ErrorAction SilentlyContinue)) {
        Show-Output "Starting Pageant."
        Start-Process -FilePath "${Pageant}" -ArgumentList @("-forcepincache", "-certauthpromptingoff", "CAPI:$($Config.Thumbprint)")
    }

    & "${GitUsrBin}\ssh-add.exe" -l *> $null
    # 2 = cannot connect to the agent
    if ($LASTEXITCODE -eq 2) {
        # A socket file may be left over from the previous session.
        Stop-Bridge
        New-Item -ItemType Directory -Path (Split-Path $SmartCardSSHSocket) -Force | Out-Null
        New-Item -ItemType Directory -Path "${SmartCardSSHDir}" -Force | Out-Null
        Show-Output "Starting ssh-pageant at `"${SmartCardSSHSocket}`"."
        $BridgeOutput = & "${GitUsrBin}\ssh-pageant.exe" -r -a $env:SSH_AUTH_SOCK -s | Out-String
        if ($LASTEXITCODE -ne 0 -or $BridgeOutput -notmatch "SSH_PAGEANT_PID=(\d+)") {
            Show-Output "Starting ssh-pageant failed: ${BridgeOutput}" -ForegroundColor Red
            exit 1
        }
        Set-Content -Path "${BridgePidPath}" -Value $Matches[1] -Encoding ASCII
    }

    # Pageant may take a moment to start and load the certificate.
    $Loaded = $false
    for ($i = 0; $i -lt 20; $i++) {
        if (Test-KeyLoaded) {
            $Loaded = $true
            break
        }
        Start-Sleep -Milliseconds 500
    }
    if (-not $Loaded) {
        Show-Output "The key was not found in Pageant." -ForegroundColor Red
        Show-Output "If Pageant was already running, a new instance cannot add keys to it."
        Show-Output "Run the script with -Restart, or add the certificate with `"Add CAPI Cert`" in the tray menu of Pageant."
        exit 1
    }

    Show-Output "Testing the key. Enter the PIN of the TPM virtual smart card if asked." -ForegroundColor Cyan
    Show-Output "The TPM locks out the card after too many wrong attempts."
    $TestPath = "${SmartCardSSHDir}\test.txt"
    Set-Content -Path "${TestPath}" -Value "Test signature by Start-SmartCardSSHAgent.ps1" -Encoding ASCII
    try {
        # With a public key, ssh-keygen signs using the agent.
        & "${GitUsrBin}\ssh-keygen.exe" -Y sign -n "smartcard-ssh-test" -f (ConvertTo-MsysPath $Config.PublicKeyPath) (ConvertTo-MsysPath $TestPath) *> $null
        $SignResult = $LASTEXITCODE
    } finally {
        Remove-Item -Path "${TestPath}", "${TestPath}.sig" -ErrorAction SilentlyContinue
    }
    if ($SignResult -ne 0) {
        Show-Output "The test signature failed. Check the PIN and try again with -Restart." -ForegroundColor Red
        exit 1
    }
    Show-Output "The key is ready for use." -ForegroundColor Green
} finally {
    $env:SSH_AUTH_SOCK = $OldAuthSock
}
