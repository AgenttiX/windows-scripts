param(
    [Parameter(Mandatory=$true)] [string] $ExportPath,
    [string] $Server = "localhost"
)

Write-Host "Virtual machines on $Server"
Get-VM -ComputerName $Server

Write-Host "Replication status"
Get-VMReplication -ComputerName $Server

$vms = Get-VMReplication -ComputerName $Server -ReplicationMode Primary | Get-VM -Name {$_.VMName} | Where-Object {$_.State -eq "off" }
echo "Virtual machines to be exported"
$vms

$vms | Export-VM -Path $ExportPath
