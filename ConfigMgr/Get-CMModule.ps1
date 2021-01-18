<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.182
	 Created on:   	1/18/2021 12:59 PM
	 Created by:   	Christian Damberg
	 Organization: 	Cygate AB
	 Filename:     	Get-CMModule.ps1
	===========================================================================
	.DESCRIPTION
		Connects to ConfigMgr powershell on client with Console installed.
#>

function Get-CMModule
{
	[CmdletBinding()]
	param ()
	Try
	{
		Write-Verbose "Attempting to import SCCM Module"
		Import-Module (Join-Path $(Split-Path $ENV:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1) -Verbose:$false
		Write-Verbose "Successfully imported the SCCM Module"
	}
	Catch
	{
		Throw "Failure to import SCCM Cmdlets."
	}
}

Get-CMModule
