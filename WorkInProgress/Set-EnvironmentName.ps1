$domainname = 'sodra.com'
$EnvironmentName = 'SodraOU'

$strName = $env:COMPUTERNAME
$strNamesearch = "$($strName)$"
$strFilter = "(&(objectCategory=Computer)(samAccountName=$strNamesearch))"

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.Filter = $strFilter

$objPath = $objSearcher.FindOne()
$objDetails = $objPath.GetDirectoryEntry()
$objDetails.RefreshCache("canonicalName")

$EnvVarOU =  ($objDetails.canonicalname -replace "/$($strName)") -replace "$domainname/" 

if ($EnvVarOU) 
{
    [System.Environment]::SetEnvironmentVariable("$EnvironmentName","$EnvVarOU",[System.EnvironmentVariableTarget]::Machine)
}