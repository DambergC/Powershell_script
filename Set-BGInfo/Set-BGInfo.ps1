<#
.Synopsis
   Configure BGinfo based on OU

.DESCRIPTION
   When running the script it check´s where the server is located in the AD and extract the path to configure BGinfo for the right environment

.EXAMPLE
   Set-BGInfo.ps1
   Basic run

.EXAMPLE
   Set-BGInfo.ps1 -FirstRun
   When running the script for the first time it will create/update registry value for run and will generate a new run.bat with the 
   values to autorun bginfo when a user log on the server/client.

.EXAMPLE
   Set-BGInfo.ps1 -Reset
   Removes registry run value if BGinfo is removed or BGinfo folder is moved.

.NOTES
   Filename: Set-BGInfo.ps1
   Author: Christian Damberg
   Website: https://www.damberg.org
   Email: christian@damberg.org
   Modified date: 2022-04-04
   Version 1.0 - First release
    
   For the script to work you need to have configured BGinfo for your environment.
   Downloadlink and instructions for BGInfo https://docs.microsoft.com/en-us/sysinternals/downloads/bginfo
        
   You need to change the value for domainname in the script.
   
   The logfile for the script is by default located in the same folder as the script. If you want the logfile at another place 
   you can change it in the script.
   
   You must create all of your configfiles for bginfo with logos and colors and save them in the same folder as the script.
   
   Update the xml-file with what configfile you want to use with which OU-name.            

   To use this script you must have OU-name connected to something that explain environment or something unique to keep servers apart.

   Exempel on OU-structure
   - Servers
        - 2012 r2
            - Production
            - Acceptance
            - Test
        - 2016
            - Production
            - Acceptance
            - test
        - 2019
            - Production
            - Acceptance
            - Test
   
#>

    #------------------------------------------------#
    # Parameters

    Param(
    [Parameter(Mandatory=$false)]
    [Switch]$Firstrun,
    [Parameter(Mandatory=$false)]
    [Switch]$Reset
    )

    #------------------------------------------------#
    # Functions in script
       
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
          "$Stamp $Level $Message" | Out-File -Encoding utf8 $logfile -Append
          }

    #------------------------------------------------#
    # Variables to the script

    [string]$DomainName = "corp.viamonstra.com"
    [string]$registrypath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
    [String]$logfile = "$PSScriptRoot\BGinfo.log"
    
    #------------------------------------------------#
    # Firstrun check 
      
    if($Firstrun -eq $true)
    {
                Write-Log -Level INFO -Message 'Delete run.bat if exist'
                Remove-Item $PSScriptRoot\run.bat -Force -ErrorAction SilentlyContinue
                write-log -Level INFO -Message 'New run.bat created'
                New-Item "$PSScriptRoot\run.bat" -ItemType File -Value "@echo off"
                Add-Content "$PSScriptRoot\run.bat" ""
                Add-Content "$PSScriptRoot\run.bat" "cd $PSScriptRoot"
                Add-Content "$PSScriptRoot\run.bat" "PowerShell.exe -ExecutionPolicy Bypass -File $PSScriptRoot\set-bginfo.ps1"
                write-log -Level INFO -Message 'Added content to run.bat'

        #Registry run check for BGinfo run
        $RegistryRunVaule = Get-ItemProperty -Path $registrypath -Name BGinfo -ErrorAction SilentlyContinue 

        if ($RegistryRunVaule -eq $null)
            {
                write-host "RegistryRun missing value, adding it to Registry" -ForegroundColor Green
                write-log -Level INFO -Message 'Missing registrysetting for run'
                New-ItemProperty -Path $registrypath -Name BGInfo -PropertyType string -Value "$PSScriptRoot\run.bat"
                write-log -Level INFO -Message 'Registry value added'
            }
        else
            {
                Write-host "Registry for run exist but will be updated" -ForegroundColor green
                write-log -Level INFO -Message 'Registryvalue exist but will be updated'
                New-ItemProperty -Path $registrypath -Name BGInfo -PropertyType string -Value "$PSScriptRoot\run.bat" -Force
            }

    }

    if($reset -eq $true)

    {
        
        Remove-ItemProperty -Path $registrypath -Name BGINFO -ErrorAction SilentlyContinue
        write-host 'Settings in Registry deleted' -ForegroundColor Red
        write-log -Level INFO -Message 'Registryvalue removed for running bginfo config at logon'
        write-log -Level INFO -Message 'Exist script'
        exit
    }

    #------------------------------------------------#
    # query to verify OU-path

    $strName = $env:COMPUTERNAME
    $strNamesearch = "$($strName)$"
    $strFilter = "(&(objectCategory=Computer)(samAccountName=$strNamesearch))"
    
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.Filter = $strFilter
    
    $objPath = $objSearcher.FindOne()
    $objDetails = $objPath.GetDirectoryEntry()
    $objDetails.RefreshCache("canonicalName")
    
    $OUPath =  ($objDetails.canonicalname -replace "/$($strName)") -replace "$domainname/"
    Write-Log -Level INFO "Find the following value when checking OU path for $env:COMPUTERNAME : $OUPath"
    
    #------------------------------------------------#    
    # Compare OU-path with variable and run BGInfo 

    # Check XML after which configfile to use
    
    [XML]$xmlfile = Get-Content .\BGinfo.xml
   
    write-log -Level INFO -Message 'Checking XML-file'
    foreach ($item in $xmlfile.bginfo.setup.name)
        {
        
            if($OUPath -like "*$item*")
            {
                $ExtractedValue = $item
            }
           }
        
        $data = $xmlfile.BGInfo.Setup | Where-Object {$_.name -contains $ExtractedValue}
   
   
    # configfile to use
    $bginfoconfig = $data.config 
    Write-log -Level INFO -Message "The script will use $bginfoconfig when running"


    # Construct executecommand
    $bgInfoExecutePath = "$PSScriptRoot\Bginfo64.exe /i$PSScriptRoot\$bgInfoConfig /timer:0 /silent /NOLICPROMPT"
    

    # Run BGinfo with BGinfo config
    write-log -Level INFO -Message "Running BGinfo with $bgInfoExecutePath"
    Invoke-Expression $bgInfoExecutePath
    
