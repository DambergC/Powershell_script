<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.182
	 Created on:   	1/18/2021 1:08 PM
	 Created by:   	Christian Damberg
	 Organization: 	Cygate AB
	 Filename:     	Set-MW-Update-On-Collection.ps1
	===========================================================================
	.DESCRIPTION
		Script reads mw.csv and configure/update maintencewindows on specific Collections
#>

$data = @()
$Collection = @()

#install CM-module
Get-CMModule -Verbose

#Get SiteCode
$SiteCode = Get-PSDrive -PSProvider CMSITE

# Set Location to SCCM
Set-Location "$($SiteCode.Name):" -Verbose


ForEach ($Coll in $data = import-csv C:\scripts\mw.csv)
{
	
	$MWs = (Get-CMMaintenanceWindow -CollectionId $Coll.Collid)
	foreach ($MW in $MWs)
	{
		
		Remove-CMMaintenanceWindow -Name $MW.name -CollectionID $Coll.Collid -Force -Verbose
	}
	
}


foreach ($item in $data = import-csv C:\scripts\mw.csv)
{
	
	
	$Schedule = New-CMSchedule -Nonrecurring -Start $item.start -end $item.end
	$Collection = Get-CMDeviceCollection -id $item.Collid
	New-CMMaintenanceWindow -CollectionID $Collection.CollectionID -Schedule $Schedule -Name $item.name -ErrorAction Ignore -ApplyTo SoftwareUpdatesOnly -Verbose
	
}

