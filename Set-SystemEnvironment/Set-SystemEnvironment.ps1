<#
.SYNOPSIS
  Add System Environment
.DESCRIPTION
  Add System Environment to use with MEMCM Hardware Inventory to populate collection and distrbute application
.PARAMETER
  DomainName - The name of your domain to be excluded after search to create record to write as System Environment
  EnvironmentName - What you want to name the variabel to be inventoried by MEMCM in Hardware Inventory
.OUTPUTS
  Log file stored in C:\Windows\Logs\SystemEnvironment.log>
.NOTES
  Version:        1.0
  Author:         Christian Damberg
  Creation Date:  15/3/2022
  Purpose/Change: Initial script development
  
.EXAMPLE
  Set-SystemEnvironment.ps1 -DomainName viamonstra.com -EnvironmentName ServerOU
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
#$ErrorActionPreference = "SilentlyContinue"

Param(
    [string]$DomainName,
    [string]$EnvironmentName
)


#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = "1.0"

#-----------------------------------------------------------[Functions & Params]------------------------------------------------------------

# Static params
$logfolder = "c:\windows\logs\"
$logfile = "$logfolder\Set-SystemEnvironment.log"



# Function to write to logfile
Function Write-Log {
  [CmdletBinding()]
  Param(
  [Parameter(Mandatory=$False)]
  [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
  [String]
  $Level = "INFO",

  [Parameter(Mandatory=$True)]
  [string]
  $Message
  )

  $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
  $Line = "$Stamp $Level $Message"
  Add-Content $logfile -Value $Line
  }

#-----------------------------------------------------------[Execution]------------------------------------------------------------


Write-log -Level INFO -Message 'Starting query about OU path' 
$strName = $env:COMPUTERNAME
$strNamesearch = "$($strName)$"
$strFilter = "(&(objectCategory=Computer)(samAccountName=$strNamesearch))"

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.Filter = $strFilter

$objPath = $objSearcher.FindOne()
$objDetails = $objPath.GetDirectoryEntry()
$objDetails.RefreshCache("canonicalName")

$EnvVarOU =  ($objDetails.canonicalname -replace "/$($strName)") -replace "$domainname/" 

Write-log -Level INFO "System Environment updated with $EnvVarOU and $EnvironmentName"

if ($EnvVarOU) 
{
   [System.Environment]::SetEnvironmentVariable("$EnvironmentName","$EnvVarOU",[System.EnvironmentVariableTarget]::Machine)
}
