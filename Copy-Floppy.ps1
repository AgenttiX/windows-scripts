. ".\utils.ps1"

$FloppyFolder = "./floppy"
if (-not (Test-Path "${FloppyFolder}")) {
    New-Item -Path "${FloppyFolder}" -ItemType "directory"
}

while ($true) {
    while ($true) {
        $Name = Read-Host "Please insert a new floppy and give a name for it"
        $FloppyPath = "${FloppyFolder}\${Name}"
        if (Test-Path "${FloppyPath}") {
            Show-Output "This folder already exists. Please use another name."
        } elseif (Test-Path "${Name}" -IsValid) {
            break
        } else {
            Show-Output "Invalid name. Please don't use special characters etc."
        }
    }

    Show-Output "Copying the floppy."
    Copy-Item -Path "A:\" -Destination "${FloppyPath}" -Recurse
    if (-not $?) {
        Show-Output "Copying the floppy failed. Please check the output folder to see which files were copied."
        continue
    }

    # $Format = Get-YesNo -Question "Floppy copied. Do you want to format the floppy?"
    # if ($Format) {
    #     Format-Volume -DriveLetter "A" -FileSystem "FAT" -Full -Force
    # }
    # Show-Output "Floppy formatted. Please remove the floppy."
    Show-Output "Floppy copied. Please remove the floppy."
}
