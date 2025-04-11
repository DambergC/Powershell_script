<#
    .SYNOPSIS
    Imports the Configuration Manager PowerShell module.

    .DESCRIPTION
    Connects to ConfigMgr PowerShell on a client with Console installed.
    Can optionally set the PowerShell location to the specified site code.

    .PARAMETER SiteCode
    Optional. The Configuration Manager site code to set as the current location.

    .PARAMETER NoSiteSwitch
    Optional. When specified with SiteCode, prevents automatically switching to the site.

    .EXAMPLE
    Get-CMModule
    # Imports the ConfigMgr module.

    .EXAMPLE
    Get-CMModule -SiteCode "P01"
    # Imports the ConfigMgr module and sets the location to P01:

    .EXAMPLE
    Get-CMModule -Verbose
    # Imports the ConfigMgr module with verbose output.

    .NOTES
    Created with: SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.182
    Created on:   1/18/2021 12:59 PM
    Created by:   Christian Damberg
    Organization: Cygate AB
    Updated on:   2025-04-11
    Filename:     Get-CMModule.ps1
#>

function Get-CMModule {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false, HelpMessage="Configuration Manager site code")]
        [ValidateNotNullOrEmpty()]
        [string]$SiteCode,
        
        [Parameter(Mandatory=$false, HelpMessage="Prevents switching to the site code")]
        [switch]$NoSiteSwitch
    )
    
    Begin {
        Write-Verbose "Starting ConfigMgr module import process"
    }
    
    Process {
        Try {
            Write-Verbose "Attempting to import SCCM Module"
            
            # Check if SMS_ADMIN_UI_PATH exists
            if (-not $ENV:SMS_ADMIN_UI_PATH) {
                Throw "Configuration Manager console is not installed or SMS_ADMIN_UI_PATH environment variable is not set."
            }
            
            $modulePath = Join-Path $(Split-Path $ENV:SMS_ADMIN_UI_PATH) "ConfigurationManager.psd1"
            
            # Check if the module file exists
            if (-not (Test-Path $modulePath)) {
                Throw "ConfigurationManager module not found at path: $modulePath"
            }
            
            # Check if module is already loaded
            if (Get-Module ConfigurationManager) {
                Write-Verbose "ConfigurationManager module is already loaded"
            } else {
                Import-Module $modulePath -Verbose:$false
                Write-Verbose "Successfully imported the SCCM Module"
            }
            
            # If SiteCode is provided and NoSiteSwitch isn't specified, set the location to the site
            if ($SiteCode -and -not $NoSiteSwitch) {
                $CMSite = $SiteCode + ":"
                Write-Verbose "Setting location to $CMSite"
                Set-Location $CMSite -ErrorAction Stop
                Write-Verbose "Successfully set location to $CMSite"
            }
            
            # Return true to indicate success
            return $true
        }
        Catch {
            $errorMessage = $_.Exception.Message
            Write-Error "Failed to import SCCM Cmdlets: $errorMessage"
            Throw "Failure to import SCCM Cmdlets: $errorMessage"
        }
    }
    
    End {
        Write-Verbose "ConfigMgr module import process completed"
    }
}

# If script is not being dot-sourced, call the function
if ($MyInvocation.InvocationName -ne '.') {
    Get-CMModule
}