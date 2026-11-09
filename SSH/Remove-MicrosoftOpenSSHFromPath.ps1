<#
.SYNOPSIS
    Remove Microsoft OpenSSH from the machine and user PATH,
    so that "ssh" keeps resolving to the OpenSSH of Git for Windows.
.DESCRIPTION
    Microsoft OpenSSH can be installed alongside Git for Windows for programs that require it,
    such as Claude Desktop, which does not accept the MSYS-based OpenSSH of Git for Windows.
    Claude Desktop finds Microsoft OpenSSH in System32 or Program Files without PATH,
    so the PATH entries can be removed safely.

    Installing or updating the OpenSSH Windows capability, the Win32-OpenSSH MSI,
    or a Windows feature update may add the entries back, so run this script again after those.

    The removed entries are:
    - %SystemRoot%\System32\OpenSSH
    - %ProgramFiles%\OpenSSH

    All other entries are kept as they are, and the registry value type
    (REG_EXPAND_SZ or REG_SZ) is preserved, so that entries such as %SystemRoot% keep working.
    The old values are backed up to %TEMP% before writing.
    Changing the machine PATH requires administrator privileges.
.PARAMETER Scope
    Which PATH variables to process. Defaults to both.
.EXAMPLE
    .\Remove-MicrosoftOpenSSHFromPath.ps1 -WhatIf
.EXAMPLE
    .\Remove-MicrosoftOpenSSHFromPath.ps1
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet("Machine", "User")]
    [string[]]$Scope = @("Machine", "User")
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"
. "${PSScriptRoot}\..\Utils.ps1"

function ConvertTo-NormalizedPath {
    <#
    .SYNOPSIS
        Expand environment variables, remove quotes and trailing separators, and convert to lowercase for comparison.
    #>
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Path
    )
    $Expanded = [Environment]::ExpandEnvironmentVariables($Path.Trim().Trim('"'))
    return $Expanded.Replace("/", "\").TrimEnd("\").ToLowerInvariant()
}

function Get-OpenSSHDirectory {
    <#
    .SYNOPSIS
        Get the normalized directories of Microsoft OpenSSH that should not be in PATH.
    #>
    [OutputType([string[]])]
    param()
    $Dirs = @("${env:SystemRoot}\System32\OpenSSH", "${env:ProgramFiles}\OpenSSH")
    # A 32-bit PowerShell would see the redirected paths instead.
    if ($env:ProgramW6432) {
        $Dirs += "${env:ProgramW6432}\OpenSSH"
    }
    return $Dirs | ForEach-Object { ConvertTo-NormalizedPath -Path $_ } | Select-Object -Unique
}

function Get-FilteredPath {
    <#
    .SYNOPSIS
        Split a PATH value and separate the Microsoft OpenSSH entries from the others.
    #>
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value,
        [Parameter(Mandatory = $true)][string[]]$ExcludedDirs
    )
    $Kept = New-Object System.Collections.Generic.List[string]
    $Removed = New-Object System.Collections.Generic.List[string]
    foreach ($Entry in $Value.Split(";")) {
        if ($Entry.Trim() -and ($ExcludedDirs -contains (ConvertTo-NormalizedPath -Path $Entry))) {
            $Removed.Add($Entry)
        } else {
            $Kept.Add($Entry)
        }
    }
    return [PSCustomObject]@{
        Value   = $Kept -join ";"
        Removed = $Removed.ToArray()
    }
}

function Send-EnvironmentChange {
    <#
    .SYNOPSIS
        Notify running programs, such as Explorer, that the environment variables have changed.
        Programs started from them afterwards get the new PATH.
    #>
    if (-not ("Win32.EnvironmentNotifier" -as [type])) {
        Add-Type -Namespace "Win32" -Name "EnvironmentNotifier" -MemberDefinition @"
[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
public static extern IntPtr SendMessageTimeout(
    IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);
"@
    }
    $HWND_BROADCAST = [IntPtr]0xffff
    $WM_SETTINGCHANGE = 0x1a
    $SMTO_ABORTIFHUNG = 0x2
    $Result = [UIntPtr]::Zero
    [void][Win32.EnvironmentNotifier]::SendMessageTimeout(
        $HWND_BROADCAST, $WM_SETTINGCHANGE, [UIntPtr]::Zero, "Environment", $SMTO_ABORTIFHUNG, 5000, [ref]$Result)
}

function Update-PathVariable {
    <#
    .SYNOPSIS
        Remove the Microsoft OpenSSH entries from the PATH of the given scope.
    .OUTPUTS
        Whether the PATH was changed.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)][ValidateSet("Machine", "User")][string]$Scope,
        [Parameter(Mandatory = $true)][string[]]$ExcludedDirs
    )
    if ($Scope -eq "Machine") {
        $Root = [Microsoft.Win32.Registry]::LocalMachine
        $SubKey = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
    } else {
        $Root = [Microsoft.Win32.Registry]::CurrentUser
        $SubKey = "Environment"
    }

    $Key = $Root.OpenSubKey($SubKey, $false)
    try {
        if (-not ($Key.GetValueNames() -contains "Path")) {
            Show-Information "${Scope} PATH: not set."
            return $false
        }
        # Read the value without expanding %VARIABLES%, so that they are written back unchanged.
        $OldValue = [string]$Key.GetValue("Path", "", [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        $Kind = $Key.GetValueKind("Path")
    } finally {
        $Key.Close()
    }

    $Filtered = Get-FilteredPath -Value $OldValue -ExcludedDirs $ExcludedDirs
    if ($Filtered.Removed.Count -eq 0) {
        Show-Information "${Scope} PATH: no Microsoft OpenSSH entries." -ForegroundColor Green
        return $false
    }
    foreach ($Entry in $Filtered.Removed) {
        Show-Information "${Scope} PATH: found `"${Entry}`"." -ForegroundColor Yellow
    }
    if (-not $PSCmdlet.ShouldProcess("${Scope} PATH", "Remove $($Filtered.Removed.Count) Microsoft OpenSSH entries")) {
        return $false
    }
    if (($Scope -eq "Machine") -and (-not (Test-Admin))) {
        throw "Changing the machine PATH requires administrator privileges. Run this script in an elevated PowerShell."
    }

    $Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $BackupPath = Join-Path $env:TEMP "path-backup_${Scope}_${Timestamp}.txt"
    Set-Content -Path $BackupPath -Value $OldValue -Encoding UTF8
    Show-Information "${Scope} PATH: backed up the old value to `"${BackupPath}`"."

    $Key = $Root.OpenSubKey($SubKey, $true)
    try {
        $Key.SetValue("Path", $Filtered.Value, $Kind)
    } finally {
        $Key.Close()
    }
    Show-Information "${Scope} PATH: removed $($Filtered.Removed.Count) entries." -ForegroundColor Green
    return $true
}

$ExcludedDirs = Get-OpenSSHDirectory
$Changed = $false
foreach ($S in $Scope) {
    if (Update-PathVariable -Scope $S -ExcludedDirs $ExcludedDirs) {
        $Changed = $true
    }
}

if ($Changed) {
    Send-EnvironmentChange
    Show-Output "Restart terminals and other programs to make them use the new PATH."
}

# Report which ssh a new process would find, using the PATH that new processes get.
$NewPath = @(
    [Environment]::GetEnvironmentVariable("Path", "Machine"),
    [Environment]::GetEnvironmentVariable("Path", "User")
) -join ";"
$FirstSSH = $NewPath.Split(";") |
    Where-Object { $_.Trim() } |
    ForEach-Object { Join-Path ([Environment]::ExpandEnvironmentVariables($_.Trim().Trim('"'))) "ssh.exe" } |
    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
    Select-Object -First 1
if ($FirstSSH) {
    Show-Output "New processes will use: ${FirstSSH}"
} else {
    Show-Output "No ssh.exe was found in PATH." -ForegroundColor Yellow
}
foreach ($Dir in @("${env:SystemRoot}\System32\OpenSSH", "${env:ProgramFiles}\OpenSSH")) {
    $Exe = Join-Path $Dir "ssh.exe"
    if (Test-Path -LiteralPath $Exe -PathType Leaf) {
        Show-Output "Microsoft OpenSSH, available to programs that look for it outside PATH: ${Exe}"
    }
}
