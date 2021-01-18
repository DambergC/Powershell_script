<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.182
	 Created on:   	1/17/2021 1:50 PM
	 Created by:   	Christian Damberg
	 Organization: 	Cygate AB
	 Filename:     	Get-Participant.ps1
	===========================================================================
	.DESCRIPTION
		Collect name of participant and the email and generate a particioant.csv
		which is used to create the xml for Cisco TMS.

		Supports multi numbers of participant.
#>


# Path to whre the csv-file should be generated 
$csvfile = 'c:\dv\participant.csv'

# If file exist and it already contains data this command will clear the file
Clear-Content $csvfile

# Define the csv-file columns
class CsvRow {
	[string]${Participant}
	[string]${Email}
}

# Ask how many participants
$numbers_of_participants = Read-host "hur många deltagare"

# foreach-loop to collect name and email about the participants
$array = foreach ($index in 1 .. $numbers_of_participants)
{
	
	$participant = Read-host "Deltagare"
	$email = read-host "Epost"
	
	# Collect the input in an hash-table
	$hash = @{
		"Participant" = $participant
		"Email"	      = $email
	}
	
	# Send the data from hash-table to variable $newRow
	$newRow = New-Object PsObject -Property $hash
	
	#Export date $newrow to the csv-file
	Export-Csv $csvfile -inputobject $newrow -append -Force -NoTypeInformation
	
}




