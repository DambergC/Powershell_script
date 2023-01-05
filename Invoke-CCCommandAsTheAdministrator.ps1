<#
.SYNOPSIS
    Invoke-CCCommandAsTheAdministrator is designed to run a command with the
    elevated permissions of The Administrator
.DESCRIPTION
    When running Invoke-CCCommandAsTheAdministrator, or using its alias of sudo,
    the command identified will be executed in another PowerShell process
    which will run as The Administrator.
        
    You will be asked for permission to use the elevated context by the UAC.
.PARAMETER Command
	This identifies the command to be executed with elevated permissions.
    There are some special options:
        !   Execute the previous command without waiting for the new process
            to complete before returning control.  The new powershell process
            does not automatically exit.
        !!  Displays the history of commands and will execute the selected 
            command without waiting for the new process to complete before 
            returning control.  The new powershell process does not 
            automatically exit.
.PARAMETER NoExit
	Aliases: noe
	When this switch is used the elevated shell does not close after completing
    the command.
.PARAMETER WaitForCompletion
	Aliases: wfc
	This switch causes the calling script to block until the elevated shell
    completes the command and exits - either automatically or manually if 
    called with the NoExit switch.
.PARAMETER PowerShellEdition
	Valid values: Default, Desktop, Core
    Default Value: "Default"
	This parameter defines which shell to use.  If set to Default the same 
    shell as the calling shell will be used.
.PARAMETER CommandLine
	This parameter is designed to collect the rest of the command line to 
    enable the Invoke-CCCommandAsTheAdministrator command to be used to launch
    an elevated shell and run a specific command.
.EXAMPLE
	Invoke-CCCommandAsTheAdministrator -Command Start-Service MSSQLServer -NoExit -WaitForCompletion -UseWindowsPowerShell
.EXAMPLE
    Invoke-CCCommandAsTheAdministrator !
    Execute the previous command as The Administrator using the same shell as the
    calling environments and do not exit the new Powershell process.
.EXAMPLE
    sudo Start-Service MSSQLServer -NoExit -WaitForCompletion -PowerShellEdition Core -Verbose
    will run PowerShell Core to start the SQL Server service.  The elevated shell will
    not close and the calling shell will block until the elevated shell is closed
    manually.

    The Verbose message will be:
    VERBOSE: Running Core edition to execute Start-Service MSSQLServer, The calling code will wait until the shell
    closes.. The shell needs to be manually closed.
.NOTES
    Command Alias: "sudo" (Linux 'Super User Do')
    Confirm Impact: Medium
    Supports Positional Binding
#>
function Invoke-CCCommandAsTheAdministrator
{
	[CmdletBinding()]
	[Alias("sudo")]
	Param
	(
		[Parameter(Position = 0)]
		[String]
		$Command,
		[Alias("noe")]
		[Switch]
		$NoExit,
		[Alias("wfc")]
		[Switch]
		$WaitForCompletion,
		[ValidateSet("Default", "Desktop", "Core")]
		[String]
		$PowerShellEdition = "Default",
		[parameter(ValueFromRemainingArguments = $true)]
		[String[]]
		$CommandLine
	)
	
	Begin
	{
		if ($PowerShellEdition -eq "Default")
		{
			$PowerShellEdition = $PSVersionTable["PSEdition"]
		}
		$parsedCommandLine = ""
		switch ($Command)
		{
			"!" { $parsedCommandLine = (Get-History | Select-Object -Last 1).CommandLine }
			"!!" { $parsedCommandLine = (Get-History | Sort-Object -Property iD -Descending | Select-Object -Property CommandLine | Out-GridView -OutputMode Single -Title "Select Command").CommandLine }
			default { $parsedCommandLine = ($Command + " " + ($CommandLine -join " ")).Trim() }
		}
		
		if ([String]::IsNullOrWhiteSpace($parsedCommandLine)) # BM: Processing Parsed Command Line
		{
			$parsedCommandLine = 'Write-Host "Administrative Shell launched from sudo"' #TODO: workaround for pwsh not responding to -NoExit with no command.
			$NoExit = $true
		}
		
		if ($NoExit)
		{
			$noExitValue = "-NoExit"
			$shellExitMessage = "needs to be manually closed"
		}
		else
		{
			$noExitValue = ""
			$shellExitMessage = "will automatically close"
		}
		if ($WaitForCompletion)
		{
			$waitMessage = "wait until the shell closes."
		}
		else
		{
			$waitMessage = "continue after launching the shell"
		}
		
		[String]$toString = ("Running {0} edition to execute {1}, The calling code will {2}. The shell {3}." -f $PowerShellEdition, $parsedCommandLine, $waitMessage, $shellExitMessage)
		Write-Verbose $toString
	}
	
	Process
	{
		switch ($PowerShellEdition)
		{
			"Core" { Start-Process -FilePath pwsh.exe -ArgumentList ($noExitValue + " -Command " + $parsedCommandLine) -Wait:$WaitForCompletion -Verb runas }
			"Desktop" { Start-Process -FilePath powershell.exe -ArgumentList ($noExitValue + " " + $parsedCommandLine) -Wait:$WaitForCompletion -Verb runas }
		}
	}
}