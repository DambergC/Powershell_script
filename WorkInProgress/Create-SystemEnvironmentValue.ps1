$strName = $env:COMPUTERNAME
$strNamesearch = "$($strName)$"
$strFilter = "(&(objectCategory=Computer)(samAccountName=$strNamesearch))"

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.Filter = $strFilter

$objPath = $objSearcher.FindOne()
$objDetails = $objPath.GetDirectoryEntry()
$objDetails.RefreshCache("canonicalName")

$EnvVarOU =  ($objDetails.canonicalname -replace "/$($strName)") -replace "sodra.com/" 


if ($EnvVarOU) 

{
[System.Environment]::SetEnvironmentVariable("SodraOU","$EnvVarOU",[System.EnvironmentVariableTarget]::Machine)
}

