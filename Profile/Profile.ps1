# This file gets loaded by both PowerShell and PowerShell ISE

# function Enable-SSHAgent {
#     . "${env:ProgramFiles}\Git\cmd\start-ssh-agent.cmd"
# }

# Import the Chocolatey Profile that contains the necessary code to enable
# tab-completions to function for `choco`.
# Be aware that if you are missing these lines from your profile, tab completion
# for `choco` will not function.
# See https://ch0.co/tab-completion for details.
$ChocolateyProfile = "${env:ChocolateyInstall}\helpers\chocolateyProfile.psm1"
if (Test-Path("${ChocolateyProfile}")) {
    Import-Module "${ChocolateyProfile}"
}

$SSHAgentStarted = $false
if (Get-Process | Where {$_.Name -eq "ssh-agent"}) {} else {
    ssh-agent > "${Env:USERPROFILE}\.ssh-agent-info" | Out-Null
    $SSHAgentStarted = $true
}
if ($SSHAgentStarted -or (-not [Environment]::GetEnvironmentVariable("SSH_AUTH_SOCK", "User"))) {
    $SSHAgentInfo = Get-Content -Path "${Env:USERPROFILE}\.ssh-agent-info"
    $Env:SSH_AUTH_SOCK = $SSHAgentInfo.split("[`n=;]")[1]
    $Env:SSH_AGENT_PID = $SSHAgentInfo.split("[`n=;]")[5]
}
if ($SSHAgentStarted) {
    & "${Env:USERPROFILE}\Git\private-scripts\ssh\Setup-Agent.ps1" | Out-Null
}

if ($env:ComputerName -eq "agx-z2e-win") {
    function Set-Monitors($Name) {
        $FilePath = "${PSScriptRoot}\DisplayConfig\${Name}.xml"
        if (Test-Path $FilePath) {
            Import-Module DisplayConfig
        } else {
            Write-Information "The monitor configuration was not found at `"${FilePath}`"".
        }
    }
}
