<#
.SYNOPSIS
    Signs files with Microsoft SignTool using Microsoft Artifact Signing.
.DESCRIPTION
    Requires the Windows SDK SignTool and the Microsoft Artifact Signing Client Tools,
    which can be installed with "winget install Microsoft.Azure.ArtifactSigningClientTools".
    The signing account is configured in a metadata file.
    Copy metadata-template.json to metadata.json and fill in the values.
    You must be authenticated to Azure, e.g. with "az login".
.PARAMETER FilePath
    Path(s) to the file(s) to sign.
.PARAMETER MetadataPath
    Path to the Artifact Signing metadata file.
    Defaults to metadata.json in the directory of this script.
.PARAMETER SignToolPath
    Path to signtool.exe. Defaults to the x64 SignTool of the newest installed Windows SDK.
.PARAMETER DlibPath
    Path to Azure.CodeSigning.Dlib.dll. Defaults to the location used by the winget package.
.PARAMETER TimestampServer
    URL of the RFC 3161 timestamp server.
.LINK
    https://learn.microsoft.com/en-us/azure/artifact-signing/how-to-signing-integrations
#>
param(
    [Parameter(Position=0, Mandatory=$true)][string[]]$FilePath,
    [string]$MetadataPath = (Join-Path "${PSScriptRoot}" "metadata.json"),
    [string]$SignToolPath,
    [string]$DlibPath = (Join-Path "${env:LOCALAPPDATA}" "Microsoft\MicrosoftArtifactSigningClientTools\Azure.CodeSigning.Dlib.dll"),
    [string]$TimestampServer = "http://timestamp.acs.microsoft.com"
)

. "${PSScriptRoot}\..\Utils.ps1"
Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

function Find-SignTool {
    <#
    .SYNOPSIS
        Find the x64 signtool.exe of the newest installed Windows SDK.
    #>
    [OutputType([string])]
    param()
    $SdkBinPath = Join-Path "${env:ProgramFiles(x86)}" "Windows Kits\10\bin"
    if (-not (Test-Path "${SdkBinPath}")) {
        return $null
    }
    # The Dlib is 64-bit, so the 64-bit SignTool must be used.
    $SignTool = Get-ChildItem -Path "${SdkBinPath}" -Directory |
        Where-Object { $_.Name -as [version] } |
        Sort-Object { [version]$_.Name } -Descending |
        ForEach-Object { Join-Path $_.FullName "x64\signtool.exe" } |
        Where-Object { Test-Path $_ } |
        Select-Object -First 1
    return $SignTool
}

if (-not $SignToolPath) {
    $SignToolPath = Find-SignTool
    if (-not $SignToolPath) {
        throw "SignTool was not found. Please install the Windows SDK or provide -SignToolPath."
    }
}
foreach ($Path in @($SignToolPath, $DlibPath, $MetadataPath)) {
    if (-not (Test-Path "${Path}" -PathType Leaf)) {
        throw "File not found: ${Path}"
    }
}
[string[]]$ResolvedFiles = @(foreach ($Path in $FilePath) {
    (Resolve-Path "${Path}").Path
})

Show-Output "Using SignTool: ${SignToolPath}"
Show-Output "Using Dlib: ${DlibPath}"
Show-Output "Using metadata: ${MetadataPath}"

& "${SignToolPath}" sign /v /debug /fd SHA256 /tr "${TimestampServer}" /td SHA256 /dlib "${DlibPath}" /dmdf "${MetadataPath}" $ResolvedFiles
if ($LASTEXITCODE -ne 0) {
    throw "SignTool failed with exit code ${LASTEXITCODE}."
}
Show-Output "Signing completed." -ForegroundColor Green
