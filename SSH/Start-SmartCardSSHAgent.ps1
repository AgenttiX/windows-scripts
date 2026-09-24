<#
.SYNOPSIS
    Start a dedicated SSH agent for the key on the TPM virtual smart card
.DESCRIPTION
    Starts the ssh-agent of Git for Windows with a fixed socket path,
    so that it can be configured with IdentityAgent in the SSH configuration,
    and also used by programs that do not have SSH_AUTH_SOCK set, such as IDEs.
    The agent is allowed to load only the OpenSC PKCS#11 module.
    The PIN is asked once when the key is added to the agent.
    After that the key can be used without prompts until the agent is stopped, the key lifetime expires,
    or the user logs out.

    Note that while the key is loaded, any process of the current user can use it through the agent.
    The key cannot be copied from the TPM, however.
.PARAMETER Lifetime
    Maximum lifetime of the key in the agent in seconds. 0 means no limit.
.PARAMETER Restart
    Stop an already running agent before starting a new one
.PARAMETER Stop
    Only stop the agent
#>

param(
    [ValidateRange(0, [int]::MaxValue)][int]$Lifetime = 0,
    [switch]$Restart,
    [switch]$Stop
)

Set-StrictMode -Version 3.0
. "${PSScriptRoot}\..\Utils.ps1"
. "${PSScriptRoot}\SmartCardUtils.ps1"

$AgentPidPath = "${SmartCardSSHDir}\agent.pid"

function Get-AgentStatus {
    <#
    .SYNOPSIS
        Get the exit code of ssh-add -l: 0 = has keys, 1 = no keys, 2 = the agent is not running
    #>
    [OutputType([int])]
    param()
    & "${GitUsrBin}\ssh-add.exe" -l *> $null
    return $LASTEXITCODE
}

function Stop-Agent {
    if (Test-Path "${AgentPidPath}") {
        $env:SSH_AGENT_PID = (Get-Content -Path "${AgentPidPath}" -Raw).Trim()
        & "${GitUsrBin}\ssh-agent.exe" -k *> $null
        Remove-Item -Path "${AgentPidPath}"
        Remove-Item Env:\SSH_AGENT_PID
    }
    if (Test-Path "${SmartCardSSHSocket}") {
        Remove-Item -Path "${SmartCardSSHSocket}"
    }
}

if (Test-Admin) {
    Show-Output "Run this script as a normal user, not elevated." -ForegroundColor Red
    exit 1
}
if (-not (Test-OpenSC)) {
    exit 1
}
$Config = Get-SmartCardSSHConfig
$PublicKey = ((Get-Content -Path $Config.PublicKeyPath -Raw).Trim() -split " ")[0..1] -join " "

$OldAuthSock = $env:SSH_AUTH_SOCK
$env:SSH_AUTH_SOCK = ConvertTo-MsysPath $SmartCardSSHSocket
try {
    $Status = Get-AgentStatus
    if ($Stop -or $Restart) {
        Show-Output "Stopping the agent."
        Stop-Agent
        $Status = 2
        if ($Stop) {
            exit 0
        }
    }

    if ($Status -eq 0) {
        $LoadedKeys = @(& "${GitUsrBin}\ssh-add.exe" -L)
        if ($LoadedKeys | Where-Object { $_.StartsWith($PublicKey) }) {
            Show-Output "The agent is already running with the key loaded." -ForegroundColor Green
            exit 0
        }
    }

    # The OpenSC configuration is re-generated to ignore any readers that have been added after the key was created.
    Set-OpenSCConfig -ReaderName $Config.ReaderName
    if ($Status -eq 2) {
        # A socket file may be left over from the previous session.
        Stop-Agent
        New-Item -ItemType Directory -Path (Split-Path $SmartCardSSHSocket) -Force | Out-Null
        Show-Output "Starting the agent at `"${SmartCardSSHSocket}`"."
        # The PKCS#11 helper of the agent inherits the environment of the agent.
        $OldConf = $env:OPENSC_CONF
        $env:OPENSC_CONF = $SmartCardSSHOpenSCConf
        try {
            $AgentOutput = & "${GitUsrBin}\ssh-agent.exe" -a $env:SSH_AUTH_SOCK -P (ConvertTo-MsysPath $OpenSCModule) | Out-String
        } finally {
            $env:OPENSC_CONF = $OldConf
        }
        if ($LASTEXITCODE -ne 0 -or $AgentOutput -notmatch "SSH_AGENT_PID=(\d+)") {
            Show-Output "Starting the agent failed: ${AgentOutput}" -ForegroundColor Red
            exit 1
        }
        Set-Content -Path "${AgentPidPath}" -Value $Matches[1] -Encoding ASCII
    }

    # OpenSSH attempts to log in to every token with the PIN, so it must not be sent to other cards.
    if (-not (Test-SingleToken)) {
        Show-Output "Not adding the key to avoid sending the PIN to other cards. Edit `"${SmartCardSSHOpenSCConf}`" or remove the other cards." -ForegroundColor Red
        exit 1
    }

    Show-Output "Adding the key to the agent. Enter the PIN of the TPM virtual smart card." -ForegroundColor Cyan
    Show-Output "The TPM locks out the card after too many wrong attempts."
    $AddArgs = @("-s", (ConvertTo-MsysPath $OpenSCModule))
    if ($Lifetime -gt 0) {
        $AddArgs = @("-t", "${Lifetime}") + $AddArgs
    }
    & "${GitUsrBin}\ssh-add.exe" @AddArgs
    if ($LASTEXITCODE -ne 0) {
        Show-Output "Adding the key failed." -ForegroundColor Red
        exit 1
    }
    $LoadedKeys = @(& "${GitUsrBin}\ssh-add.exe" -L)
    if (-not ($LoadedKeys | Where-Object { $_.StartsWith($PublicKey) })) {
        Show-Output "The key `"$($Config.PublicKeyPath)`" was not found in the agent. The loaded keys are:" -ForegroundColor Red
        $LoadedKeys | ForEach-Object { Show-Output "  $_" }
        exit 1
    }
    Show-Output "The key is loaded to the agent." -ForegroundColor Green
} finally {
    $env:SSH_AUTH_SOCK = $OldAuthSock
}
