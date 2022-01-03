<#
.Synopsis
   Crypted Passwordfile
.DESCRIPTION
   The script ask for path where to save a crypted passwordfile to be 
   used in automation where elevated credential are needed,
.EXAMPLE
   New-CryptedPasswordFile -Filepath .\CryptedPasswordfilename.txt
.NOTES

		===========================================================================
		Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2021 v5.8.183
		Created on:   	1/27/2021 10:35 AM
		Updated on:		  1/27/2021 10:35 AM
		Created by:   	Christian Damberg, Sebastian Thörngren
		Organization: 	Cygate AB
		Filename:     	New-CryptedPasswordFile.ps1
		===========================================================================
#>
function New-CryptedPasswordFile
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Filepath

    )

    Begin
    {
    }
    Process
    {
    
    $credential = Get-Credential
    $credential.Password | ConvertFrom-SecureString | Set-Content $Filepath
    }
    End
    {
    }
}
New-CryptedPasswordFile