<#
.SYNOPSIS
    Utility methods for scripts
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidAssignmentToAutomaticVariable", "", Justification="Overriding only for old PowerShell versions")]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseDeclaredVarsMoreThanAssignments", "LogPath", Justification="Used in scripts")]
param()

# Compatibility for old PowerShell versions
if($PSVersionTable.PSVersion.Major -lt 3) {
    # Configure PSScriptRoot variable for old PowerShell versions
    # https://stackoverflow.com/a/8406946/
    # https://stackoverflow.com/a/5466355/
    $invocation = (Get-Variable MyInvocation).Value
    $PSScriptRoot = Split-Path $invocation.MyCommand.Path

    # Implement missing features

    function Invoke-WebRequest {
        <#
        .SYNOPSIS
            Download a file from the web
        .LINK
            https://stackoverflow.com/a/43544903
        #>
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidOverwritingBuiltInCmdlets", "", Justification="Overriding only for old PowerShell versions")]
        param(
            [Parameter(Mandatory=$true)][string]$Uri,
            [Parameter(Mandatory=$true)][string]$OutFile
        )
        $WebClient_Obj = New-Object System.Net.WebClient
        # TODO: this fails for some reason
        $WebClient_Obj.DownloadFile($Uri, $OutFile)
    }

    function Expand-Archive {
        <#
        .SYNOPSIS
            Zip file decompression
        .DESCRIPTION
            TODO: test that this works
        .LINK
            https://stackoverflow.com/a/54687028
        #>
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidOverwritingBuiltInCmdlets", "", Justification="Overriding only for old PowerShell versions")]
        param(
            [Parameter(Mandatory=$true)][string]$Path,
            [Parameter(Mandatory=$true)][string]$DestinationPath
        )
        $shell_ComObject = New-Object -ComObject shell.application
        $zip_file = $shell_ComObject.namespace($Path)
        $folder = $shell_ComObject.namespace($DestinationPath)
        $folder.Copyhere($zip_file.items())
    }
}

# This has to be after the compatibility section so that $PSScriptRoot is defined
# $RepoPath = Split-Path $PSScriptRoot -Parent
$RepoPath = $PSScriptRoot

# Create sub-directories of the repo so that the scripts don't have to create them.
New-Item -Path "$RepoPath" -Name "Downloads" -ItemType "directory" -Force | Out-Null
New-Item -Path "$RepoPath" -Name "Logs" -ItemType "directory" -Force | Out-Null
$Downloads = "${RepoPath}\Downloads"
$LogPath = "${RepoPath}\Logs"

# Define some useful paths
$CommonDesktopPath = [Environment]::GetFolderPath("CommonDesktopDirectory")
$DesktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
$CommonStartMenuPath = "${env:PROGRAMDATA}\Microsoft\Windows\Start Menu\Programs"
$UserStartMenuPath = "${env:APPDATA}\Microsoft\Windows\Start Menu\Programs"
$StartPath = Get-Location

# Define some useful variables
$UserDir = ([System.IO.DirectoryInfo]"${env:UserProfile}").FullName
$RepoInUserDir = ([System.IO.DirectoryInfo]"${RepoPath}").FullName.StartsWith($UserDir)

# This will be filled by Get-InstalledSoftware
$InstalledSoftware = $null

# Force Invoke-WebRequest to use TLS 1.2 and TLS 1.3 instead of the insecure TLS 1.0, which is the default.
# https://stackoverflow.com/a/41618979
if ([Net.SecurityProtocolType]::Tls13) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
} else {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}


function Add-ScriptShortcuts {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Justification="Plural on purpose")]
    param()

    if ($RepoInUserDir) {
        $BasePath = $DesktopPath
    } else {
        if (! (Test-Admin)) {
            Show-Output "Cannot create script shortcuts for all users, as the script is not elevated (yet)."
            return
        }
        $BasePath = $CommonDesktopPath
    }
    $InstallShortcutPath = "${BasePath}\Installer.lnk"
    $MaintenanceShortcutPath = "${BasePath}\Maintenance.lnk"
    $WindowsUpdateShortcutPath = "${BasePath}\Windows Update.lnk"
    $ReportShortcutPath = "${BasePath}\Create IT report.lnk"
    $RepoShortcutPath = "${BasePath}\windows-scripts.lnk"

    # if(!(Test-Path $InstallShortcutPath)) {
    New-Shortcut -Path $InstallShortcutPath -TargetPath "powershell" -Arguments "-File ${RepoPath}\Install-Software.ps1" -WorkingDirectory "${RepoPath}" -IconLocation "shell32.dll,21" | Out-Null
    # }
    # if(!(Test-Path $MaintenanceShortcutPath)) {
    New-Shortcut -Path $MaintenanceShortcutPath -TargetPath "powershell" -Arguments "-File ${RepoPath}\Maintenance.ps1" -WorkingDirectory "${RepoPath}" -IconLocation "shell32.dll,80" | Out-Null
    # }
    New-Shortcut -Path $ReportShortcutPath -TargetPath "powershell" -Arguments "-File ${RepoPath}\Report.ps1" -WorkingDirectory "${RepoPath}" -IconLocation "shell32.dll,1" | Out-Null
    New-Shortcut -Path $RepoShortcutPath -TargetPath "${RepoPath}" | Out-Null

    $WindowsVersion = [System.Environment]::OSVersion.Version.Major
    if ($WindowsVersion -ge 8) {
        if ((Get-CimInstance Win32_OperatingSystem).Caption -Match "Windows 10") {
            $WindowsUpdateTargetPath = "ms-settings:windowsupdate-action"
        } else {
            $WindowsUpdateTargetPath = "ms-settings:windowsupdate"
        }
        New-Shortcut -Path $WindowsUpdateShortcutPath -TargetPath "${env:windir}\explorer.exe" -Arguments "${WindowsUpdateTargetPath}" -IconLocation "shell32.dll,46" | Out-Null
    }
}

function Clear-Path {
    <#
    .SYNOPSIS
        Clear the path for e.g. unpacking a program.
    .DESCRIPTION
        Return values: 0 not cleared, 1 cleared, 2 no need
    #>
    # [Parameter(Mandatory=$true)]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory=$true)][string]$Path
    )
    if (Test-Path "${Path}") {
        $Decision = $Host.UI.PromptForChoice("The path `"${Path}`" already exists.", "Shall I clear and overwrite it?", ("y", "n"), 0)
        if ($Decision -eq 0) {
            Remove-Item -Path "${Path}" -Recurse
            return 1
        }
        return 0
    }
    return 2
}

function Elevate {
    <#
    .SYNOPSIS
        Elevate the current process to admin privileges.
    .LINK
        https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/start-process?view=powershell-7.2#example-5--start-powershell-as-an-administrator
    #>
    param(
        [Parameter(Mandatory=$true)][string]$command
    )
    if (! (Test-Admin))  {
        if ($Elevated)
        {
            Show-Output "Elevation did not work."
        }
        else {
            Show-Output "This script requires admin access. Elevating."
            Show-Output "${command}"
            $ArgumentList = ('-NoProfile -NoExit -Command "cd {0}; {1} -Elevated"' -f ($pwd, $command))
            # Use newer PowerShell if available.
            if (Test-CommandExists "pwsh") {$shell = "pwsh"} else {$shell = "powershell"}
            # Use Windows Terminal if available
            # if (Test-CommandExists "wt") {
            #     # https://stackoverflow.com/a/63163528
            #     # This has problems in parsing the semicolon
            # Start-Process -FilePath "wt" -Verb RunAs -ArgumentList ('new-tab {0} {1}' -f ($shell, $ArgumentList))
            # } else {
            Start-Process -FilePath "$shell" -Verb RunAs -ArgumentList "${ArgumentList}"
            # }
            Show-Output "The script has been started in another window. You can close this window now." -ForegroundColor Green
        }
    Exit
    }
}

function Find-First {
    <#
    .SYNOPSIS
        Find the first file that matches the given filter
    .LINK
        https://stackoverflow.com/a/1500148/
    #>
    [OutputType([System.IO.FileInfo])]
    param(
        [string]$Filter,
        [string]$Path
    )
    return @(Get-ChildItem -Filter "${Filter}" -Path "${Path}")[0]
}

function Get-CertificateInfo {
    <#
    .SYNOPSIS
        Get basic info of a certificate as a string
    #>
    [OutputType([string])]
    param([Parameter(Mandatory=$true)][System.Security.Cryptography.X509Certificates.X509Certificate]$Certificate)
    $NotBefore = $Certificate.NotBefore.ToUniversalTime().ToString("yyyy-MM-dd")
    $NotAfter = $Certificate.NotAfter.ToUniversalTime().ToString("yyyy-MM-dd")
    $Algorithm = $Certificate.SignatureAlgorithm.FriendlyName
    return "$($Certificate.Subject) (${NotBefore} - ${NotAfter}, ${Algorithm}, " `
        + "Currently valid: $($Certificate.Verify()), Thumbprint: $($Certificate.Thumbprint))"
}

function Get-InstallBitness {
    [OutputType([string])]
    param(
        [string]$x86 = "x86",
        [string]$x86_64 = "x86_64",
        [switch]$Quiet = $false
    )
    if ([System.Environment]::Is64BitOperatingSystem) {
        if (-not $Quiet) {
            Show-Information "64-bit operating system detected. Installing 64-bit version."
        }
        return $x86_64
    }
    if (-not $Quiet) {
        Show-Information "32-bit operating system detected. Installing 32-bit version."
    }
    return $x86
}

function Get-InstalledSoftware {
    <#
    .SYNOPSIS
        Get a list of installed software.
    .NOTES
        This uses Win32_Product, which is abysmally slow. Find a better method if possible.
        Unfortunately, Get-Product does not work in PowerShell 7.
    .LINK
        https://gregramsey.net/2012/02/20/win32_product-is-evil/
    #>
    [OutputType([System.Array])]
    param()
    if ($null -eq $script:InstalledSoftware) {
        Show-Information "Getting the list of installed software. This may take a while."
        $script:InstalledSoftware = Get-CimInstance -ClassName Win32_Product
    }
    return $script:InstalledSoftware
}

function Get-IsDomainJoined {
    [OutputType([bool])]
    param()
    return (Get-CimInstance -ClassName Win32_ComputerSystem).PartOfDomain
}

function Get-IsHeadlessServer {
    <#
    .SYNOPSIS
        Test whether the computer is a headless server (e.g. Windows Server Core or Hyper-V Server)
    .LINK
        https://serverfault.com/a/529131
     #>
    [OutputType([bool])]
    param()
    # return Get-IsServer -and (-not (Get-WindowsFeature -Name 'Server-Gui-Shell').InstallState -eq 2)
    return Test-Path "${env:windir}\explorer.exe"
}

function Get-IsServer {
    [OutputType([bool])]
    param()
    return (Get-CimInstance -Class Win32_OperatingSystem).ProductType -ne 1
}

function Get-IsVirtualMachine {
    [OutputType([bool])]
    param()
    return Get-IsVirtualBoxMachine -or ((Get-CimInstance Win32_ComputerSystem).Model -contains "Virtual")
}

function Get-IsVirtualBoxMachine {
    [OutputType([bool])]
    param()
    return (
        (Test-Path "${env:ProgramFiles}\Oracle\VirtualBox Guest Additions") -or
        ((Get-CimInstance Win32_ComputerSystem).Model -eq "VirtualBox")
    )
}

function Get-YesNo {
    [OutputType([bool])]
    param(
        [Parameter(Mandatory=$true)][string]$Question
    )
    while ($true) {
        $Confirmation = Read-Host "${Question} [y/n]"
        if ($Confirmation -eq "y") {
            return $True
        }
        if ($Confirmation -eq "n") {
            return $False
        }
    }
}

function Install-Chocolatey {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingInvokeExpression", "", Justification="Chocolatey installation requires Invoke-Expression")]
    param(
        [switch]$Force
    )
    if($Force -Or (-Not (Test-CommandExists "choco"))) {
        Show-Output "Installing the Chocolatey package manager."
        Set-ExecutionPolicy Bypass -Scope Process -Force
        # This is already configured globally above
        # [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString("https://chocolatey.org/install.ps1"))
    }
}

function Install-Executable {
    param(
        [string]$Name,
        [string]$Path,
        [switch]$BypassAuthenticode = $false,
        [string]$MD5,
        [string]$SHA256
    )
    # Validate the file path
    try {
        $File = Get-Item "${Path}" -ErrorAction Stop
    } catch {
        if ($Path.StartsWith("${RepoPath}")) {
            Show-Output -ForegroundColor Red "${Name} installer was not found at ${Path}. Do you have the network drive mounted?"
            Show-Output -ForegroundColor Red "It could be that your computer does not have the necessary group policies applied. Applying. You will need to reboot for the changes to become effective."
            gpupdate /force
        }
        throw [System.IO.FileNotFoundException] "${Name} installer was not found at `"${Path}`"."
    }
    # Validate checksum
    Test-Checksum -Path $Path -MD5 $MD5 -SHA256 $SHA256
    # Validate Authenticode
    if (-not $BypassAuthenticode) {
        Test-AuthenticodeSignature -FilePath "${Path}"
    }
    # Install the executable
    Show-Output "Installing ${Name}"
    if ($File.Extension -eq ".msi") {
        $Process = Start-Process -NoNewWindow -Wait -PassThru "msiexec" -ArgumentList "/i","${Path}"
    } else {
        $Process = Start-Process -NoNewWindow -Wait -PassThru "${Path}"
    }
    $ExitCode = $Process.ExitCode
    if ($ExitCode) {
        Show-Output -ForegroundColor Red "${Name} installation returned non-zero exit code ${ExitCode}. Perhaps the installation failed?"
    }
}

function Install-FromUri {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingBrokenHashAlgorithms", "", Justification="Legacy MD5 support is on purpose for backwards compatibility")]
    param(
        [Parameter(mandatory=$true)][string]$Name,
        [Parameter(mandatory=$true)][string]$Uri,
        [Parameter(mandatory=$false)][string]$Filename,
        [Parameter(mandatory=$false)][string]$UnzipFolderName,
        [Parameter(mandatory=$false)][string]$UnzippedFilePath,
        [Parameter(mandatory=$false)][string]$MD5,
        [Parameter(mandatory=$false)][string]$SHA256,
        [switch]$BypassAuthenticode = $false
    )
    # Check whether $Uri is a remote resource or a path from e.g. a network drive.
    $IsRemote = $Uri.Contains("://")
    if ($IsRemote) {
        if (($Uri.Contains("http://") -or ($Uri.Contains("ftp://"))) -and (-not ($PSBoundParameters.ContainsKey("SHA256") -or $PSBoundParameters.ContainsKey("MD5")))) {
            throw [System.Security.SecurityException] "Downloading over an unencrypted connection without providing a hash for checking file integrity is not supported for security."
        }
        $Path = "${Downloads}\${Filename}"
        if (-not $PSBoundParameters.ContainsKey("Filename")) {
            throw [System.ArgumentException] "A filename must be provided when downloading a remote file."
        }
    } else {
        $Path = $Uri
    }

    # Test if the file already exists and matches the given hash
    $FileAlreadyOK = $false
    if (Test-Path $Path) {
        if ($IsRemote) {
            $DownloadText1 = "is already downloaded and "
            $DownloadText2 = " Skipping download."
            $DownloadText3 = "already downloaded"
            $DownloadText4 = " The previous download may have been interrupted. Downloading again."
        } else {
            $DownloadText1 = ""
            $DownloadText2 = ""
            $DownloadText3 = "found"
            $DownloadText4 = " Please request a healthy copy from the software developer."
        }
        if ($PSBoundParameters.ContainsKey("SHA256")) {
            Show-Output "Verifying SHA256 checksum for `"${Path}`"."
            $FileHash = (Get-FileHash -Path "${Path}" -Algorithm "SHA256").Hash
            if ($FileHash -eq $SHA256) {
                $FileAlreadyOK = $true
                Show-Output "The file `"${Path}`" ${DownloadText1}has the correct SHA-256 hash ${SHA256}.${DownloadText2}"
            } else {
                Show-Output -ForegroundColor Red "The file `"${Path}` was ${DownloadText3}, but has an incorrect SHA-256 hash. " `
                    + "(Expected ${SHA256}, got ${FileHash}.)${DownloadText4}"
                if (-not $IsRemote) {
                    throw [System.Security.SecurityException]
                }
            }
        } elseif ($PSBoundParameters.ContainsKey("MD5")) {
            Show-Output "Verifying MD5 checksum for `"${Path}`"."
            $FileHash = (Get-FileHash -Path "${Path}" -Algorithm "MD5").Hash
            if ($FileHash -eq $MD5) {
                $FileAlreadyOK = $true
                Show-Output "The file `"${Path}`" ${DownloadText1}has the correct MD5 hash ${MD5}.${DownloadText2}"
            } else {
                Show-Output -ForegroundColor Red "The file `"${Path}` was ${DownloadText3}, but has an incorrect MD5 hash. " `
                    + "(Expected ${MD5}, got ${FileHash}.)${DownloadText4}"
                if (-not $IsRemote) {
                    throw [System.Security.SecurityException]
                }
            }
        }
    } elseif (-not $IsRemote) {
        throw [System.IO.FileNotFoundException] "The path `"${Path}`" seems not to exist on the system, nor does it seem to be a remote uri."
    }

    # Download the file
    if ($IsRemote -and (-not $FileAlreadyOK)) {
        Show-Output "Downloading ${Name}"
        Invoke-WebRequestFast -Uri "${Uri}" -OutFile "${Path}"
    }

    # Get the file object
    try {
        $File = Get-Item "${Path}" -ErrorAction Stop
    } catch {
        throw [System.IO.FileNotFoundException] "The file was not found at ${Path}"
    }

    # Verify the file checksum if not already verified
    if (-not $FileAlreadyOK) {
        Test-Checksum -Path $Path -MD5 $MD5 -SHA256 $SHA256
    }

    # Process the downloaded file
    if ($File.Extension -eq ".zip") {
        # Some zip files have a directory already in them.
        # if (-not $PSBoundParameters.ContainsKey("UnzipFolderName")) {
        #     Show-Output -ForegroundColor Red "UnzipFolderName was not provided for a zip file."
        #     return 1
        # }
        if ($PSBoundParameters.ContainsKey("UnzipFolderName")) {
            $UnzipPath = "${Downloads}\${UnzipFolderName}"
            if ((Test-Path $UnzipPath) -and (-not (Clear-Path $UnzipPath))) {
                throw [System.IO.IOException] "Cannot extract to the non-empty folder `"${UnzipPath}`"."
            }
        } else {
            $UnzipPath = $Downloads
        }
        Show-Output "Extracting ${Name}"
        # The -Force is necessary for overwriting an existing .exe
        Expand-Archive -LiteralPath "${Path}" -DestinationPath "${UnzipPath}" -Force
        if ($PSBoundParameters.ContainsKey("UnzippedFilePath")) {
            $ExecutablePath = "${Downloads}\${UnzipFolderName}\${UnzippedFilePath}"
        } else {
            Show-Output "UnzippedFilePath was not provided for the zip file."
            return
        }
    } else {
        $ExecutablePath = $Path
    }
    Install-Executable -Name "${Name}" -Path "${ExecutablePath}" -BypassAuthenticode:$BypassAuthenticode
}

function Install-Geekbench {
    <#
    .SYNOPSIS
        Install the Geekbench performance test benchmarks
    .LINK
        https://www.geekbench.com/download/windows/
    .LINK
        https://www.geekbench.com/legacy/
    #>
    param(
        [string[]]$GeekbenchVersions = @("6.2.1", "5.5.1", "4.4.4", "3.4.4", "2.4.3")
    )
    foreach ($Version in $GeekbenchVersions) {
        Show-Output "Downloading Geekbench ${Version}"
        $Filename = "Geekbench-$Version-WindowsSetup.exe"
        $Url = "https://cdn.geekbench.com/$Filename"
        Invoke-WebRequest -Uri "$Url" -OutFile "${Downloads}\${Filename}"
    }
    foreach ($Version in $GeekbenchVersions) {
        Show-Output "Installing Geekbench ${Version}"
        Start-Process -NoNewWindow -Wait "${Downloads}\Geekbench-${Version}-WindowsSetup.exe"
    }
}

function Install-PTS {
    param(
        [string]$PTS_version = "10.8.4",
        [bool]$Force = $False
    )
    $PTS = "${Env:SystemDrive}\phoronix-test-suite\phoronix-test-suite.bat"
    if ((-not (Test-Path "$PTS")) -or $Force) {
        Install-FromUri
            -Name "Phoronix Test Suite (PTS)"
            -Uri "https://github.com/phoronix-test-suite/phoronix-test-suite/archive/v${PTS_version}.zip"
            -Filename "phoronix-test-suite-${PTS_version}.zip"

        # The installation script needs to be executed in its directory
        try {
            Set-Location "${Downloads}\phoronix-test-suite-${PTS_version}\phoronix-test-suite-${PTS_version}"
            Show-Output "Installing Phoronix Test Suite (PTS)."
            Show-Output "You may get prompts asking whether you accept the EULA and whether to send anonymous usage statistics."
            Show-Output "Please select yes to the former and preferably to the latter as well."
            & ".\install.bat"
        } finally{
            Set-Location "$StartPath"
        }
        if (-not (Test-Path "$PTS")) {
            throw [System.IO.FileNotFoundException] "Phoronix Test Suite (PTS) was not found at `"${PTS}`"."
        }
        Show-Output "Phoronix Test Suite (PTS) has been installed. It is highly recommended that you log in now so that you can manage the uploaded results."
        & "$PTS" openbenchmarking-login
        Show-Output "You must now configure the batch mode according to your preferences. Some of the settings may get overwritten later by this script."
        & "$PTS" batch-setup
    }
    & "$PTS" openbenchmarking-refresh
}

function Install-Winget {
    [OutputType([bool])]
    param()
    if (Test-CommandExists "winget") {
        return $true
    }
    if (Get-AppxPackage -Name "Microsoft.DesktopAppInstaller") {
        Show-Output "App Installer seems to be installed on your system, but Winget was not found."
        return $false
    }
    if (Get-AppxPackage -Name "Microsoft.WindowsStore") {
        Show-Output "App Installer appears not to be installed. Please close this window and install it from the Windows Store. Then restart this script."
        Start-Process -Path "https://www.microsoft.com/en-us/p/app-installer/9nblggh4nns1"
        $confirmation = Show-Output "If you know what you're doing, you may also continue by writing `"force`", but some features may be disabled.".
        while ($confirmation -ne "force") {
            $confirmation = Show-Output "Close this window or write `"force`" to continue."
        }
        return $false;
    }
    Show-Output "Cannot install App Installer, as Microsoft Store appears not to be installed. This is normal on servers. Winget will not be available."
    return $false
}

function Invoke-WebRequestFast {
    <#
    .SYNOPSIS
        Invoke-WebRequest but without the progress bar that slows it down
    .LINK
        https://github.com/PowerShell/PowerShell/issues/13414
    .LINK
        https://learn.microsoft.com/en-us/virtualization/community/team-blog/2017/20171219-tar-and-curl-come-to-windows
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [Parameter(Mandatory=$true)][string]$OutFile
    )
    if (Test-CommandExists "curl.exe") {
        # The --http2 and --http3 arguments are not supported on Windows 11 22H2.
        # https://github.com/microsoft/WSL/issues/3141
        # --location is used instead of --url so that redirects are followed, which is required for GitHub downloads.
        curl.exe --location "${Uri}" --output "${OutFile}" --tlsv1.2
    } else {
        $PreviousProgressPreference = $ProgressPreference
        $ProgressPreference = "SilentlyContinue"
        Invoke-WebRequest -Uri $Uri -OutFile $OutFile
        $ProgressPreference = $PreviousProgressPreference
    }
}

function New-Junction {
    <#
    .SYNOPSIS
        Create a new directory junction (equivalent to a symlink on Linux)
    .LINK
        https://github.com/PowerShell/PowerShell/issues/621
     #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Target,
        [Parameter(Mandatory=$false)][boolean]$Backup = $true,
        [Parameter(Mandatory=$false)][string]$BackupSuffix = "-old",
        [Parameter(Mandatory=$false)][string]$BackupPath
    )
    if (Test-Path -Path "${Path}") {
        $LinkType = (Get-Item -Path "${Path}" -Force).LinkType
        if ($LinkType -or (-not $Backup)) {
            Show-Output "Removing old junction at `"${Path}`", LinkType=${LinkType}"
            # The -Recurse argument has to be specified to remove a junction.
            # This doesn't remove the actual directory or its contents. Please see the link above.
            if($PSCmdlet.ShouldProcess($Path, "Remove-Item")) {
                Remove-Item "${Path}" -Recurse
            }
        } else {
            if (-not $PSBoundParameters.ContainsKey("BackupPath")) {
                $BackupPath = "${Path}${BackupSuffix}"
            }
            Show-Output "Backing up old directory at `"${Path}`" to `"${BackupPath}`""
            if($PSCmdlet.ShouldProcess($Path, "Move-Item")) {
                Move-Item -Path "${Path}" -Destination "${BackupPath}"
            }
        }
    }
    Show-Output "Creating junction from `"${Path}`" to `"${Target}`""
    if($PSCmdlet.ShouldProcess($Path, "New-Item")) {
        return New-Item -ItemType "Junction" -Path "${Path}" -Target "${Target}"
    }
}

function New-Shortcut {
    <#
    .SYNOPSIS
        Create a .lnk shortcut
    .LINK
        https://stackoverflow.com/a/9701907
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$TargetPath,
        [string]$Arguments,
        [string]$IconLocation,
        [string]$WorkingDirectory
    )
    # Some shortcuts may also be created for bare program names,
    # in which case Test-Path may not find them, even though they are valid.
    if ([System.IO.Path]::IsPathRooted($TargetPath) -and (-not (Test-Path($TargetPath)))) {
        throw [System.IO.FileNotFoundException] "Could not find shortcut target `"${TargetPath}`""
    }

    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($Path)
    $Shortcut.TargetPath = $TargetPath
    if (! [string]::IsNullOrEmpty($WorkingDirectory)) {
        $Shortcut.WorkingDirectory = $WorkingDirectory
    }
    if (! [string]::IsNullOrEmpty($Arguments)) {
        $Shortcut.Arguments = $Arguments
    }
    if (! [string]::IsNullOrEmpty($IconLocation)) {
        $Shortcut.IconLocation = $IconLocation
    }
    if($PSCmdlet.ShouldProcess($Path, "Save")) {
        $Shortcut.Save()
    }
    return $Shortcut
}

function New-DesktopShortcut {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)][string]$TargetPath,
        [string]$Arguments,
        [string]$IconLocation,
        [string]$WorkingDirectory,
        [switch]$UserOnly
    )
    if ($UserOnly -or ($TargetPath.StartsWith($UserDir, "CurrentCultureIgnoreCase"))) {
        $Path = "${DesktopPath}\$Name.lnk"
    } else {
        $Path = "${CommonDesktopPath}\${Name}.lnk"
    }
    New-Shortcut -Path $Path -TargetPath $TargetPath -Arguments $Arguments -WorkingDirectory $WorkingDirectory
}

function New-StartMenuShortcut {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)][string]$TargetPath,
        [string]$Arguments,
        [string]$IconLocation,
        [string]$WorkingDirectory,
        [string]$UserOnly
    )
    if ($UserOnly -or ($TargetPath.StartsWith($UserDir, "CurrentCultureIgnoreCase"))) {
        $Path = "${UserStartMenuPath}\${Name}.lnk"
    } else {
        $Path = "${CommonStartMenuPath}\${Name}.lnk"
    }
    New-Shortcut -Path $Path -TargetPath $TargetPath -Arguments $Arguments -WorkingDirectory $WorkingDirectory
}

function Request-DomainConnection {
    [OutputType([bool])]
    $IsJoined = Get-IsDomainJoined
    if ($IsJoined) {
        Show-Output "Your computer seems to be a domain member. Please connect it to the domain network now. A VPN is OK but a physical connection is better."
    }
    return $IsJoined
}

function Set-RepoPermissions {
    [CmdletBinding(SupportsShouldProcess)]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Justification="Plural on purpose")]
    param()

    Import-Module Microsoft.PowerShell.Security
    if ($RepoInUserDir) {
        Show-Output "The repo is installed within the user folder. Ensuring that it's owned by the user."
        $ACL = Get-Acl -Path "${RepoPath}"
        # This does not include the domain part
        # $User = New-Object System.Security.Principal.Ntaccount([Environment]::UserName)
        $User = whoami
        $CurrentOwner = $ACL.Owner
        if ($CurrentOwner -ne $User) {
            Show-Output "The repo is owned by `"${CurrentOwner}`". Changing it to: `"${User}`""
            $ACL.SetOwner($User)
            if($PSCmdlet.ShouldProcess($RepoPath, "Set-Acl")) {
                $ACL | Set-Acl -Path "${RepoPath}"
            }
        }
    } else {
        Show-Output "The repo is installed globally. Ensuring that it's protected."
        $ACL = Get-Acl -Path "${env:ProgramFiles}"
        if($PSCmdlet.ShouldProcess($RepoPath, "Set-Acl")) {
            $ACL | Set-Acl -Path "${RepoPath}"
        }
        $RepoPathObj = (Get-Item $RepoPath)
        if (($RepoPathObj.Parent.FullName -eq "Git") -and ($RepoPathObj.Parent.Parent.FullName -eq $env:SystemDrive)) {
            Show-Output "The repo is located in the Git folder at the root of the system drive. Protecting the Git folder."
            $ACL = Get-Acl -Path $env:SystemDrive
            if($PSCmdlet.ShouldProcess($RepoPathObj, "Set-Acl")) {
                $ACL | Set-Acl -Path $RepoPathObj.Parent
            }
        }
    }
}

function Show-Information {
    <#
    .SYNOPSIS
        Write-Information with colors. Use this instead of Write-Host.
    #>
    param(
        [Parameter(Position=0, Mandatory=$true)]$MessageData,
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor
        # This is set globally
        # https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_commonparameters?view=powershell-7.2#-informationaction
        # [string]$InformationAction = "Continue"
    )
    if ($InformationPreference -eq "SilentlyContinue") {
        $CustomInformationPreference = "Continue"
    } else {
        $CustomInformationPreference = $InformationPreference
    }
    Show-Stream @PSBoundParameters -Stream "Write-Information" -InformationAction $CustomInformationPreference
}

function Show-Output {
    <#
    .SYNOPSIS
        Write-Output with colors.
    #>
    param(
        [Parameter(Position=0, Mandatory=$true)]$InputObject,
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor
    )
    Show-Stream @PSBoundParameters -Stream "Write-Output"
}

function Show-Stream {
     <#
    .SYNOPSIS
        Show text or an object on the console.
    .DESCRIPTION
        With color support but without using Write-Host.
    .LINK
        https://stackoverflow.com/a/4647985/
    #>
    # Argument passing
    # https://stackoverflow.com/a/62861781/
    param(
        [Parameter(Position=0, Mandatory=$true)][Alias("MessageData")]$InputObject,
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor,
        [string]$Stream = "Write-Output"
    )
    if ($PSBoundParameters.ContainsKey("ForegroundColor")) {
        $OldForegroundColor = [Console]::ForegroundColor
        [Console]::ForegroundColor = $ForegroundColor
    }
    if ($PSBoundParameters.ContainsKey("BackgroundColor")) {
        $OldBackgroundColor = [Console]::BackgroundColor
        [Console]::BackgroundColor = $BackgroundColor
    }
    if ($args) {
        & $Stream $InputObject $args
    } elseif ($Stream -eq "Write-Information") {
        # Write-Information does not support piping
        & $Stream $InputObject
    } else {
        $InputObject | & "${Stream}"
    }
    if ($PSBoundParameters.ContainsKey("ForegroundColor")) {
        [Console]::ForegroundColor = $OldForegroundColor
    }
    if ($PSBoundParameters.ContainsKey("BackgroundColor")) {
        [Console]::BackgroundColor = $OldBackgroundColor
    }
}

function Test-Admin {
    <#
    .SYNOPSIS
        Test whether the script is being run as an administrator
    .LINK
        https://superuser.com/questions/108207/how-to-run-a-powershell-script-as-administrator
    #>
    [OutputType([bool])]
    $CurrentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    return $CurrentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

function Test-AuthenticodeSignature {
    <#
    .SYNOPSIS
        Validate that a file has a valid Authenticode signature
    #>
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [switch]$Silent
    )
    $Signature = Get-AuthenticodeSignature -FilePath "${FilePath}"
    $Failed = $Signature.Status -ne "Valid"
    if (-not $Silent) {
        $Status = "Status: $($Signature.Status) ($($Signature.StatusMessage))"
        if ($Failed) {
            Show-Output -ForegroundColor Red "Authenticode signature verification failed for `"$($Signature.Path)`"."
            Show-Output -ForegroundColor Red "${Status}"
            Show-Output -ForegroundColor Red "This may be an indication of malicious tampering with the file."
            Show-Output -ForegroundColor Red "Please contact your system administrator or the developers of the software."
        } else {
            Show-Output "Authenticode signature verification successful for `"$($Signature.Path)`"."
            Show-Output "${Status}, Signature type: $($Signature.SignatureType)"
        }
        $Signer = $Signature.SignerCertificate
        if ($Signer) {
            Show-Output "Certificate: $(Get-CertificateInfo $Signer)"
        }
        $TimeStamper = $Signature.TimeStamperCertificate
        if ($TimeStamper) {
             Show-Output "Timestamper: $(Get-CertificateInfo $TimeStamper)"
        }
    }
    if ($Failed) {
        throw "Authenticode signature verification failed for `"$($Signature.Path)`"."
    }
}

function Test-Checksum {
    param(
        [string]$Path,
        [string]$MD5,
        [string]$SHA256
    )
    if ($SHA256) {
        Show-Output "Verifying SHA256 checksum for `"${Path}`"."
        $FileHash = (Get-FileHash -Path "${Path}" -Algorithm "SHA256").Hash
        if ($FileHash -eq $SHA256) {
            Show-Output "SHA256 checksum OK"
        } else {
            throw [System.Security.SecurityException] "The file has an invalid SHA256 checksum. Expected: ${SHA256}, got: ${FileHash}"
        }
    } elseif ($MD5) {
        Show-Output "Verifying MD5 checksum for `"${Path}`"."
        $FileHash = (Get-FileHash -Path "${Path}" -Algorithm "MD5").Hash
        if ($FileHash -eq $MD5) {
            Show-Output "MD5 checksum OK"
        } else {
            throw [System.Security.SecurityException] "The file has an invalid MD5 checksum. Expected: ${MD5}, got: ${FileHash}"
        }
    }
}

function Test-CommandExists {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Justification="Not plural")]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory=$true)][string]$command
    )
    $OldPreference = $ErrorActionPreference
    $ErrorActionPreference = "stop"
    try {
        if (Get-Command $command){
            return $true
        }
    } catch {
        return $false
    } finally {
        $ErrorActionPreference=$OldPreference
    }
}

function Test-PendingRebootAndExit {
    # The PendingReboot module gives false positives about the need to reboot.
    # if (Test-CommandExists "Install-Module") {
    #     Show-Output -ForegroundColor Cyan "Ensuring that the PowerShell reboot checker module is installed. You may now be asked whether to install the NuGet package provider. Please select yes."
    #     Install-Module -Name PendingReboot -Force
    # }
    # if ((Test-CommandExists "Test-PendingReboot") -and (Test-PendingReboot -SkipConfigurationManagerClientCheck -SkipPendingFileRenameOperationsCheck).IsRebootPending) {
    if (Test-RebootPending) {
        Show-Output -ForegroundColor Cyan "A reboot is already pending. Please close this window, reboot the computer and then run this script again."
        if (! (Get-YesNo "If you are sure you want to continue regardless, please write `"y`" and press enter.")) {
            exit 0
        }
    }
}

function Test-RebootPending {
    <#
    .SYNOPSIS
        Test whether the computer has a reboot pending. The name is chosen not to conflict with the PendingReboot PowerShell module.
    .LINK
        https://4sysops.com/archives/use-powershell-to-test-if-a-windows-server-is-pending-a-reboot/
    #>
    [OutputType([bool])]
    $PendingRebootTests = @(
        @{
            Name = "RebootPending"
            Test = { Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing" -Name "RebootPending" -ErrorAction Ignore }
            TestType = "ValueExists"
        }
        @{
            Name = "RebootRequired"
            Test = { Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" -Name "RebootRequired" -ErrorAction Ignore }
            TestType = "ValueExists"
        }
        @{
            Name = "PendingFileRenameOperations"
            Test = { Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name "PendingFileRenameOperations" -ErrorAction Ignore }
            TestType = "NonNullValue"
        }
    )
    foreach ($Test in $PendingRebootTests) {
        $result = Invoke-Command -ScriptBlock $Test.Test
        if ($Test.TestType -eq "ValueExists" -and $result) {
            return $true
        } elseif ($Test.TestType -eq "NonNullValue" -and $result -and $result.($Test.Name)) {
            return $true
        } else {
            return $false
        }
    }
}

function Update-Repo {
    <#
    .SYNOPSIS
        Update the repository if enough time has passed since the previous "git pull"
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([bool])]
    param(
        [TimeSpan]$MaxTimeSpan = (New-TimeSpan -Days 1)
    )
    if (! ($RepoInUserDir -or (Test-Admin))) {
        Show-Output "Cannot update the repo, as the script is not elevated (yet)."
        return
    }
    $FetchHeadPath = "${RepoPath}\.git\FETCH_HEAD"
    if (!(Test-Path "${FetchHeadPath}")) {
        Show-Output "The date of the previous `"git pull`" could not be determined. Updating."
        if($PSCmdlet.ShouldProcess($RepoPath, "git pull")) {
            git pull
        }
        Show-Output "You may have to restart the script to use the new version."
        # return $true
        return
    }
    $LastPullTime = (Get-Item "${FetchHeadPath}").LastWriteTime
    $TimeDifference = New-TimeSpan -Start $LastPullTime -End (Get-Date)
    if ($TimeDifference -lt $MaxTimeSpan) {
        Show-Output "Previous `"git pull`" was run on ${LastPullTime}. No need to update."
        # return $false
    } else {
        Show-Output "Previous `"git pull`" was run on ${LastPullTime}. Updating."
        if ($RepoInUserDir -and (! $Elevated)) {
            $GitConfigPath = "${env:UserProfile}\.gitconfig"
            $RepoName = Split-Path "${RepoPath}" -Leaf
            if ((! (Test-Path "${GitConfigPath}")) -or (! (Select-String "${GitConfigPath}" -Pattern "${RepoName}" -SimpleMatch))) {
                Show-Output "Git config was not found, or it did not contain the repository path. Adding ${RepoPath} as a safe directory."
                if($PSCmdlet.ShouldProcess($RepoPath, "git config")) {
                    git config --global --add safe.directory "${RepoPath}"
                }
            }
        }
        if($PSCmdlet.ShouldProcess($RepoPath, "git pull")) {
            git pull
        }
        Show-Output "You may have to restart the script to use the new version."
        # return $true
    }
}
