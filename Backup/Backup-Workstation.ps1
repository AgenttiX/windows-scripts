<#
.SYNOPSIS
    Backup the user home directories of a workstation to a network share.
.DESCRIPTION
    Copies the user profile directory tree (C:\Users by default) to an SMB network share
    with Robocopy, which is significantly faster than Copy-Item, supports multithreading
    and only copies the files that have changed since the previous backup.

    The script requires admin privileges for reading files that the current user does not
    have access to, and elevates itself automatically.
.PARAMETER Source
    The directory to be backed up.
.PARAMETER BackupRoot
    The root of the backups on the network share. The backup is created in a subdirectory
    named after the computer.
.PARAMETER ComputerName
    The name of the subdirectory to back up into. Defaults to the hostname of this computer.
.PARAMETER Destination
    The full destination path. Overrides -BackupRoot and -ComputerName.
.PARAMETER ExcludedDirectories
    Names or paths of the subdirectories to exclude from the backup.
    A bare name such as "AppData" excludes every directory with that name anywhere in the tree.
.PARAMETER ExcludedFiles
    Names or wildcard patterns of the files to exclude from the backup.
.PARAMETER Threads
    The number of Robocopy threads. Increase for network shares with a high latency.
.PARAMETER Retries
    The number of retries for failed files. Keep low, as locked files are common in user profiles.
.PARAMETER RetryWaitSeconds
    The wait time between the retries in seconds.
.PARAMETER CopyFlags
    The file properties to copy. D = data, A = attributes, T = timestamps,
    S = NTFS access control lists (ACLs), O = owner info, U = auditing info.
    Copying the ACLs requires that the destination file system supports them.
.PARAMETER NoBackupMode
    Do not use the Robocopy backup mode. Backup mode requires the "Back up files and directories"
    user right, which the administrators have by default. Without the backup mode the files
    that the current user has no access to are skipped instead of being copied.
.PARAMETER Mirror
    Delete the files from the destination that no longer exist in the source.
    This makes the backup an exact mirror, but it also means that accidentally deleted files
    are removed from the backup on the next run.
.PARAMETER ListOnly
    Only list the files that would be copied, without copying anything.
.PARAMETER Force
    Do not ask for confirmation when mirroring.
.PARAMETER Elevated
    Used internally to detect whether the script has already been elevated.
.EXAMPLE
    .\Backup-Workstation.ps1
.EXAMPLE
    .\Backup-Workstation.ps1 -ExcludedDirectories @("AppData", "Downloads") -Mirror
.LINK
    https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSReviewUnusedParameter", "Elevated", Justification="Used in utils")]
param(
    [string]$Source = "${env:SystemDrive}\Users",
    [string]$BackupRoot = "V:\IT\IT administration\Backup",
    [string]$ComputerName = "${env:ComputerName}",
    [string]$Destination,
    [string[]]$ExcludedDirectories = @(
        ".cache",
        ".cargo",
        ".lmstudio\models",
        ".matplotlib",
        ".phoronix-test-suite",
        ".ssh",
        ".vscode",
        ".vscode-oss",
        ".wakatime",
        "__pycache__",
        '$RECYCLE.BIN',
        "AppData",
        "Application Data",
        "cache",
        "Documents\PowerShell",
        "http-cache",
        "Local Settings",
        "Temporary Internet Files",
        "OneDrive",
        "OneDrive - *",
        "OneDriveTemp",
        "Searches",
        "venv",
        "${env:SystemDrive}\Users\Default",
        "${env:SystemDrive}\Users\Default User"
    ),
    [string[]]$ExcludedFiles = @(
        "*.cache",
        "*.ovpn",
        "*.tmp",
        ".bash_history",
        ".ssh-agent-info",
        ".wakatime.cfg",
        "desktop.ini",
        "hiberfil.sys",
        "NTUSER.DAT*",
        "ntuser.ini",
        "pagefile.sys",
        "swapfile.sys"
        "Thumbs.db",
        "UsrClass.dat*"
    ),
    [ValidateRange(1, 128)][int]$Threads = 16,
    [ValidateRange(0, 100)][int]$Retries = 1,
    [ValidateRange(0, 3600)][int]$RetryWaitSeconds = 5,
    [ValidatePattern("^[DATSOU]+$")][string]$CopyFlags = "DAT",
    [switch]$NoBackupMode,
    [switch]$Mirror,
    [switch]$ListOnly,
    [switch]$Force,
    [switch]$Elevated
)

Set-StrictMode -Version 3.0

. "$(Join-Path $((Get-Item "${PSScriptRoot}").Parent.FullName) "Utils.ps1")"


function Get-ElevationCommand {
    <#
    .SYNOPSIS
        Create the command for re-running this script elevated with the same arguments.
    .DESCRIPTION
        Elevation restarts the script in a new process, so the arguments have to be
        passed on explicitly. The values are single-quoted so that the elevated shell
        does not expand them.
    #>
    [OutputType([string])]
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][hashtable]$BoundParameters
    )
    $Parts = @("& '$($Path.Replace("'", "''"))'")
    foreach ($Parameter in $BoundParameters.GetEnumerator()) {
        # The Elevated switch is appended by Elevate itself.
        if ($Parameter.Key -eq "Elevated") {
            continue
        }
        $Value = $Parameter.Value
        if ($Value -is [switch]) {
            $Parts += "-$($Parameter.Key):`$$($Value.IsPresent)"
        } elseif ($Value -is [array]) {
            $Items = foreach ($Item in $Value) { "'$("${Item}".Replace("'", "''"))'" }
            $Parts += "-$($Parameter.Key) @($($Items -join ', '))"
        } else {
            $Parts += "-$($Parameter.Key) '$("${Value}".Replace("'", "''"))'"
        }
    }
    return $Parts -join " "
}

function Get-UncPath {
    <#
    .SYNOPSIS
        Convert a path on a mapped network drive into a UNC path.
    .DESCRIPTION
        Mapped network drives belong to a logon session and are not visible to elevated
        processes, as User Account Control (UAC) gives the elevated process a separate token.
        The mapping itself is stored per user, so it can be looked up from the registry
        and the drive letter replaced with the UNC path of the share.
        Paths that are not on a mapped drive are returned unchanged.
    .LINK
        https://learn.microsoft.com/en-us/troubleshoot/windows-client/networking/mapped-drives-not-available-from-elevated-command
    #>
    [OutputType([string])]
    param(
        [Parameter(Mandatory=$true)][string]$Path
    )
    if ($Path -notmatch "^([A-Za-z]):\\?(.*)$") {
        return $Path
    }
    $DriveLetter = $Matches[1]
    $RelativePath = $Matches[2]

    $RemotePath = $null
    $Drive = Get-PSDrive -Name "${DriveLetter}" -ErrorAction SilentlyContinue
    if ($Drive -and $Drive.DisplayRoot -and "$($Drive.DisplayRoot)".StartsWith("\\")) {
        $RemotePath = $Drive.DisplayRoot
    }
    if (-not $RemotePath) {
        # The drive is not mapped in this session. Look up the mapping of the user.
        $RegistryPath = "HKCU:\Network\${DriveLetter}"
        if (Test-Path "${RegistryPath}") {
            $RemotePath = (Get-ItemProperty -Path "${RegistryPath}" -Name "RemotePath" -ErrorAction SilentlyContinue).RemotePath
        }
    }
    if (-not $RemotePath) {
        return $Path
    }
    if ($RelativePath) {
        return (Join-Path "${RemotePath}" "${RelativePath}")
    }
    return $RemotePath
}

function Resolve-Destination {
    <#
    .SYNOPSIS
        Ensure that the destination directory exists and is writable.
    .DESCRIPTION
        Returns the path that is actually usable, which may be a UNC path
        if the destination is on a mapped network drive.
    #>
    [OutputType([string])]
    param(
        [Parameter(Mandatory=$true)][string]$Path
    )
    $Candidates = @("${Path}")
    $UncPath = Get-UncPath "${Path}"
    if ($UncPath -ne $Path) {
        $Candidates += "${UncPath}"
    }
    foreach ($Candidate in $Candidates) {
        try {
            if (-not (Test-Path -Path "${Candidate}")) {
                New-Item -Path "${Candidate}" -ItemType "Directory" -Force -ErrorAction Stop | Out-Null
            }
            return $Candidate
        } catch {
            Show-Output -ForegroundColor Yellow "The destination `"${Candidate}`" is not available: $($_.Exception.Message)"
        }
    }
    throw [System.IO.IOException] (
        "The backup destination `"${Path}`" could not be created or accessed. " +
        "Please check that the network share is mounted and that you have write access to it. " +
        "Note that mapped network drives are not visible to elevated processes unless " +
        "EnableLinkedConnections is enabled, in which case a UNC path has to be used instead."
    )
}

function Get-RobocopyResult {
    <#
    .SYNOPSIS
        Convert a Robocopy exit code into a human-readable description.
    .LINK
        https://learn.microsoft.com/en-us/troubleshoot/windows-server/backup-and-storage/return-codes-used-robocopy-utility
    #>
    [OutputType([string])]
    param(
        [Parameter(Mandatory=$true)][int]$ExitCode
    )
    if ($ExitCode -eq 0) {
        return "No files needed to be copied."
    }
    $Descriptions = @()
    if ($ExitCode -band 1) { $Descriptions += "Files were copied." }
    if ($ExitCode -band 2) { $Descriptions += "Extra files or directories were detected in the destination." }
    if ($ExitCode -band 4) { $Descriptions += "Mismatched files or directories were detected." }
    if ($ExitCode -band 8) { $Descriptions += "Some files or directories could not be copied." }
    if ($ExitCode -band 16) { $Descriptions += "Fatal error. No files were copied." }
    return $Descriptions -join " "
}


Elevate(Get-ElevationCommand -Path $MyInvocation.MyCommand.Definition -BoundParameters $PSBoundParameters)

if (-not $Destination) {
    $Destination = Join-Path "${BackupRoot}" "${ComputerName}"
}

if (-not (Test-Path -Path "${Source}")) {
    throw [System.IO.DirectoryNotFoundException] "The source directory `"${Source}`" was not found."
}
$Destination = Resolve-Destination "${Destination}"

if ($Mirror -and (-not $ListOnly) -and (-not $Force)) {
    Show-Output -ForegroundColor Yellow (
        "Mirroring is enabled. The files that no longer exist in `"${Source}`" " +
        "will be deleted from `"${Destination}`"."
    )
    if (-not (Get-YesNo "Do you want to continue?")) {
        exit 1
    }
}

$Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$RobocopyLogPath = "${LogPath}\Backup-Workstation_${ComputerName}_${Timestamp}.log"

$RobocopyArguments = @(
    "${Source}",
    "${Destination}",
    # Copy all subdirectories, including the empty ones
    "/E",
    # Copy the given file properties and the directory timestamps and attributes
    "/COPY:${CopyFlags}",
    "/DCOPY:DAT",
    # Do not follow junctions. Without this the legacy compatibility junctions of the user
    # profiles, e.g. "Application Data", cause the same files to be copied many times.
    "/XJ",
    # Locked files are common in user profiles, so do not stall on them for long.
    "/R:${Retries}",
    "/W:${RetryWaitSeconds}",
    # Multithreading. This is the main performance benefit over Copy-Item.
    "/MT:${Threads}",
    # Do not print the per-file progress percentage, as it would flood the log
    "/NP",
    # Print the output to both the console and the log file
    "/TEE",
    "/LOG+:${RobocopyLogPath}"
)
if ($NoBackupMode) {
    # Restartable mode
    $RobocopyArguments += "/Z"
} else {
    # Restartable mode, falling back to backup mode for the files that cannot be read otherwise.
    # Backup mode requires the backup user right, which is why this script elevates itself.
    $RobocopyArguments += "/ZB"
}
if ($Mirror) {
    $RobocopyArguments += "/PURGE"
}
if ($ListOnly) {
    $RobocopyArguments += "/L"
}
if ($ExcludedDirectories) {
    $RobocopyArguments += "/XD"
    $RobocopyArguments += $ExcludedDirectories
}
if ($ExcludedFiles) {
    $RobocopyArguments += "/XF"
    $RobocopyArguments += $ExcludedFiles
}

Show-Output "Backing up `"${Source}`" to `"${Destination}`""
Show-Output "Excluded directories: $($ExcludedDirectories -join ', ')"
Show-Output "Excluded files: $($ExcludedFiles -join ', ')"
Show-Output "Log file: `"${RobocopyLogPath}`""

$StartTime = Get-Date
& "${env:SystemRoot}\System32\robocopy.exe" @RobocopyArguments
# Robocopy uses exit codes below 8 for reporting the results of a successful run,
# so the exit code cannot be treated as an error code as such.
$ExitCode = $LASTEXITCODE
$Duration = New-TimeSpan -Start $StartTime -End (Get-Date)

Show-Output "Robocopy exit code ${ExitCode}: $(Get-RobocopyResult $ExitCode)"
Show-Output "Duration: $($Duration.ToString("hh\:mm\:ss"))"

if ($ExitCode -ge 8) {
    Show-Output -ForegroundColor Red "The backup failed or was incomplete. Please see the log at `"${RobocopyLogPath}`"."
    if (-not $NoBackupMode) {
        Show-Output -ForegroundColor Red "If the backup mode is not available on this computer, please try again with -NoBackupMode."
    }
    exit $ExitCode
}
Show-Output -ForegroundColor Green "The backup of `"${Source}`" to `"${Destination}`" is ready."
# Robocopy uses non-zero exit codes for successful runs as well,
# so the exit code is reset for the benefit of e.g. scheduled tasks.
exit 0
