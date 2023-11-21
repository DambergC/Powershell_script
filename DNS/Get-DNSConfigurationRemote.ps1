<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2023 v5.8.229
	 Created on:   	2023-10-09 01:44
	 Created by:   	Administrator
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>

Function Get-DNSServerSearchOrder
{
   <#
        .SYNOPSIS
            Attempts to retrieve the DNS server list in any active NICs with static IPs on a remote computer.

        .DESCRIPTION
            Retrieves the DNS server list for active NICs with static IPs on a remote computer, with optional alternative credentials.
            Returns a PS object for each NIC containing details of the NICs and result of the command. Optional parameter for matching specific IPs

        .PARAMETER ComputerName
            Target computer name.

        .PARAMETER FilterIPs
            A list of DNS server IPs to test if present in each NIC
            
        .PARAMETER Credential
            Optional PowerShell credential object to log on to target computer.

        .NOTES
            Name: Get-DNSServerSearchOrder

        .EXAMPLE
            Get-DNSServerSearchOrder -ComputerName 'MyComputer' -Credential (Get-Credential) -FilterIPs ('192.168.148.11','192.168.150.201') -Verbose

            Description
            -----------
            Returns DNS servers on active NICs for "MyComputer". Will prompt for credentials to log on to the target and test each NIC for the listed IP addresses.

        .EXAMPLE
            Get-DNSServerSearchOrder -ComputerName 'MyComputer' -FilterIPs ('192.168.148.11','192.168.150.201') -Verbose

            Description
            -----------
            Returns DNS servers on active NICs for "MyComputer". Will use current user credentials to log on to the target and test each NIC for the listed IP addresses.

        .EXAMPLE
            $UserID = 'DOMAIN\UserID'
            PS C:\>$PlainPassword = 'mypassword'
            PS C:\>$SecPwd = ConvertTo-SecureString $PlainPassword -AsPlainText -Force
            PS C:\>$Cred = New-Object System.Management.Automation.PSCredential ($UserID,$SecPwd)
            PS C:\>$dns = '192.168.148.11','192.168.150.201'
            
            PS C:\>Get-DNSServerSearchOrder -ComputerName 'MyComputer' -Credential $Cred -Verbose

            Description
            -----------
            Returns DNS servers on active NICs for "MyComputer" using the supplied credential object.

    #>
	[CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = "High")]
	Param (
		[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
		$ComputerName,
		[Parameter(Position = 2)]
		[String[]]$FilterIPs,
		[PSCredential]$Credential
	)
	Write-Verbose "Checking $ComputerName"
	If (Test-Connection $ComputerName -Count 1 -ErrorAction SilentlyContinue)
	{
		Write-Verbose 'Target responding to ICMP'
		If ($Credential)
		{
			Try
			{
				$Message = "Retrieving WMI object using supplied credentials: " + $Credential.UserName
				Write-Verbose $Message
				$NICs = Get-WmiObject -Class win32_networkadapterconfiguration -Credential $Credential -ComputerName $ComputerName -Filter "IPEnabled=TRUE" -ErrorAction Stop
			}
			Catch
			{
				If ($_.Exception.Message)
				{
					Return [pscustomobject]@{
						ComputerName	 = $ComputerName;
						IPAddress	     = $null;
						Description	     = $null;
						DefaultIPGateway = $null;
						DNSServers	     = $null;
						MatchFilter	     = $null;
						Result		     = $_.Exception.Message
					}
				}
				Else
				{
					Return [pscustomobject]@{
						ComputerName	 = $ComputerName;
						IPAddress	     = $null;
						Description	     = $null;
						DefaultIPGateway = $null;
						DNSServers	     = $null;
						MatchFilter	     = $null;
						Result		     = $_
					}
				}
				Write-Verbose 'Unable to retrieve WMI settings'
			}
		}
		Else
		{
			Try
			{
				Write-Verbose "Retrieving WMI object using current user credentials: $Env:USERNAME"
				$NICs = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -Filter "IPEnabled=TRUE" -ErrorAction Stop
			}
			Catch
			{
				If ($_.Exception.Message)
				{
					Return [pscustomobject]@{
						ComputerName	 = $ComputerName;
						IPAddress	     = $null;
						Description	     = $null;
						DefaultIPGateway = $null;
						DNSServers	     = $null;
						MatchFilter	     = $null;
						Result		     = $_.Exception.Message
					}
				}
				Else
				{
					Return [pscustomobject]@{
						ComputerName	 = $ComputerName;
						IPAddress	     = $null;
						Description	     = $null;
						DefaultIPGateway = $null;
						DNSServers	     = $null;
						MatchFilter	     = $null;
						Result		     = $_
					}
				}
				Write-Verbose 'Unable to retrieve WMI settings'
			}
		}
		ForEach ($NIC in $NICs)
		{
			$Script:MatchFilter = $null
			If ($FilterIPs)
			{
				$NIC.DNSServerSearchOrder | ForEach-Object {
					If ($FilterIPs -contains $_)
					{
						$FilterIPs
						$Script:MatchFilter = $true
					}
				}
			}
			[pscustomobject]@{
				ComputerName		 = $ComputerName;
				IPAddress		     = $NIC.IPAddress;
				Description		     = $NIC.Description;
				DefaultIPGateway	 = $NIC.DefaultIPGateway;
				DNSServerSearchOrder = ($NIC.DNSServerSearchOrder | Where-Object { $_ -ne $null });
				MatchFilter		     = $Script:MatchFilter;
				Result			     = 'OK'
			}
		}
	}
	Else
	{
		Write-Verbose 'No ping response'
		Return [pscustomobject]@{
			ComputerName		 = $ComputerName;
			IPAddress		     = $null;
			Description		     = $null;
			DefaultIPGateway	 = $null;
			DNSServerSearchOrder = $null;
			MatchFilter		     = $null;
			Result			     = 'No ping response'
		}
	}
	Write-Verbose 'Done'
}
