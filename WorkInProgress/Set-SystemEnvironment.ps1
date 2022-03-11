<#
.SYNOPSIS
  Add System Environment
.DESCRIPTION
  Add System Environment to use with MEMCM Hardware Inventory to populate collection and distrbute application
.PARAMETER <Parameter_Name>
    <Brief description of parameter input required. Repeat this attribute if required>
.INPUTS
  None
.OUTPUTS
  Log file stored in C:\Windows\Logs\SystemEnvironment.log>
.NOTES
  Version:        1.0
  Author:         Christian Damberg
  Creation Date:  10/3/2022
  Purpose/Change: Initial script development
  
.EXAMPLE
  Set-SystemEnvironment.ps1
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

Param(
    [string]$domainname,
    [string]$EnvironmentName
)


#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = "1.0"

#Log File Info
$sLogPath = "C:\Windows\Temp"
$sLogName = "SystemEnvironment.log"
$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

#Script variables
#$domainname = 'sodra.com'
#$EnvironmentName = 'SodraOU'

#-----------------------------------------------------------[Functions & Params]------------------------------------------------------------

Function Write-Log
{
    PARAM(
    [String]$Message,
    [int]$Severity,
    [string]$Component
    )
    Set-Location $PSScriptRoot
    $Logpath = "C:\Windows\Logs"
    $TimeZoneBias = Get-CimInstance win32_timezone
    $Date= Get-Date -Format "HH:mm:ss.fff"
    $Date2= Get-Date -Format "MM-dd-yyyy"
        "<![LOG[$Message]LOG]!><time=$([char]34)$Date$($TimeZoneBias.bias)$([char]34) date=$([char]34)$date2$([char]34) component=$([char]34)$Component$([char]34) context=$([char]34)$([char]34) type=$([char]34)$Severity$([char]34) thread=$([char]34)$([char]34) file=$([char]34)$([char]34)>"| Out-File -FilePath "$Logpath\$sLogName" -Append -NoClobber -Encoding default

}

#-----------------------------------------------------------[Execution]------------------------------------------------------------



Write-Log -Message 'Start Check'
$strName = $env:COMPUTERNAME
$strNamesearch = "$($strName)$"
$strFilter = "(&(objectCategory=Computer)(samAccountName=$strNamesearch))"

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.Filter = $strFilter

$objPath = $objSearcher.FindOne()
$objDetails = $objPath.GetDirectoryEntry()
$objDetails.RefreshCache("canonicalName")

$EnvVarOU =  ($objDetails.canonicalname -replace "/$($strName)") -replace "$domainname/" 

Write-Log -Message "$envvarOU"

if ($EnvVarOU) 
{
   # [System.Environment]::SetEnvironmentVariable("$EnvironmentName","$EnvVarOU",[System.EnvironmentVariableTarget]::Machine)
}

