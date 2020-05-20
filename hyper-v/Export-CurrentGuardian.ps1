<#
.SYNOPSIS
    Exports the Hyper-V host guardian configuration of the current computer to an xml file for migrating virtual machines with vTPMs.
    Run this on the destination host.
.PARAMETER Path
    The path to which to output the destination guardian file. Should end with .xml.
.NOTES
    Based on
    https://threadfin.com/allowing-an-additional-host-to-run-a-vm-with-virtual-tpm/
    https://gist.github.com/larsiwer/0eef34728d6697f8d5b899b1d866e573#file-update-keyprotectorforallvms-ps1
#>

param (
    [string]$Path = ".\DestinationGuardian.xml"
)

$ExportedGuardianPath = $Path

# Retrieving the current Host Guardian Service client configuration
$HgsClientConfiguration = Get-HgsClientConfiguration

If (($HgsClientConfiguration| Select -ExpandProperty Mode) -eq "HostGuardianService")
{ # The destination system is configured to use a Host Guardian Service
  # Build the URL to download the HGS's metadata.xml
  $MetadataXmlUrl = $HgsClientConfiguration.KeyProtectionServerUrl, "service/metadata/2014-07/metadata.xml" -join "/"

  # Download the metadata.xml file and save to specified file name
  Invoke-WebRequest -UseBasicParsing -Uri $MetadataXmlUrl -OutFile $ExportedGuardianPath
}
else
{ # Destination system is running in local mode
  # Get the local UntrustedGuardian from the destination machine
  $UntrustedGuardian = Get-HgsGuardian -Name UntrustedGuardian -ErrorAction SilentlyContinue

  If (!$UntrustedGuardian)
  {
    # Creating new UntrustedGuardian since it did not exist
    $UntrustedGuardian = New-HgsGuardian -Name UntrustedGuardian –GenerateCertificates
  }

  # Exporting the UntrustedGuardian to the destination path
  $UntrustedGuardian | Export-HgsGuardian -Path $ExportedGuardianPath
}
