[CmdletBinding()]
Param(   
    [parameter(Mandatory=$false, HelpMessage="Encrypt the secure app settings")]
    [switch]$Encrypt,

    [parameter(Mandatory=$false, HelpMessage="Decrypt the secure app settings")]
    [switch]$Decrypt
)

$ErrorActionPreference = "Stop"

function Get-EncryptionMode()
{
    if($Encrypt.IsPresent -and $Decrypt.IsPresent) { Write-Error "Use either -Encrypt or -Decrypt" }

    if($Decrypt.IsPresent)
    { 
        return "Decrypt"
    }
    
    return "Encrypt"
}

function Find-ConfigurationFile()
{
    $configFiles  = @(Get-ChildItem -Filter "*.exe.config")
    $configFiles += @(Get-ChildItem -Filter "App.config")

    if($configFiles.Count -gt 1) { Write-Error ("Found more than one configuration file: {0}" -f [string]::Join(", ", $configFiles.Name)) }
    if($configFiles.Count -ne 1) { Write-Error "Could not find the .config file" }

    return $configFiles.Get(0)
}

function Get-DotNetFrameworkDirectory()
{
    $([System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory())
}

function Copy-ToTempWebConfig($configurationFile)
{
    if (!(Test-Path -path ".\temp")) 
    {
        Write-Verbose ("Creating temp directory {0}" -f ".\temp")
        New-Item ".\temp" -Type Directory
    }

    Copy-Item $configurationFile "temp\Web.config"
    return Get-Item "temp\Web.config"
}

function Copy-FromTempWebConfig($configurationFile)
{
    Move-Item "temp\Web.config" $configurationFile -Force | Out-Null

    if((Get-ChildItem ".\temp").Count -eq 0)
    {
        Write-Verbose ("Removing empty temp directory {0}" -f ".\temp")
        Remove-Item ".\temp"
    }
}

function Encrypt-ConfigurationSection([string] $configurationPath, $mode){  
  $currentDirectory = (Get-Location)
  Set-Location (Get-DotNetFrameworkDirectory)

  if($mode -eq "Decrypt")
  {
    .\aspnet_regiis -pdf "secureAppSettings" "$configurationPath"
  } else
  {
    .\aspnet_regiis -pef "secureAppSettings" "$configurationPath"
  }
  Set-Location $currentDirectory
}

$mode = Get-EncryptionMode

$configurationFile = (Find-ConfigurationFile)
Write-Verbose ("{0} configuation file {1}" -f $mode, $configurationFile.FullName)

$tempFile = Copy-ToTempWebConfig $configurationFile
Write-Verbose ("Attempting to {0} {1}" -f $mode, $configurationFile.FullName)
Encrypt-ConfigurationSection $tempFile.Directory.FullName $mode
Copy-FromTempWebConfig $configurationFile