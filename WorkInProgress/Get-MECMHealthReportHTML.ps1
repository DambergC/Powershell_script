<#
    .Synopsis
    Generates an overview of client health in your Configuration Manager environment as an html email
    .DESCRIPTION
    This script dynamically builds an html report based on key client health related data from the Configuration Manager database. The script is intended to be run regularly as a scheduled task.
    .EXAMPLE
    To run as a scheduled task, use a command like the following:
    Powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "<path>\New-CMClientHealthSummaryReport.ps1"
    .NOTES
    More information can be found here: http://smsagent.wordpress.com/free-configmgr-reports/client-health-summary-report/
    The Parameters region should be updated for your environment before executing the script
#>


#######################################################################
#region Parameters
#######################################################################
# Database info
$script:dataSource = 'th-mgt02' # SQL Server name (and instance where applicable)
$script:database = 'CM_PS1' # ConfigMgr Database name

# Email params
$EmailParams = @{
  To         = 'christian.damberg@Cygate.se'
  From       = 'MECMStatus@trivselhus.se'
  Smtpserver = 'webmail.trivselhus.se'
  Subject    = "ConfigMgr Client Health Summary  |  $(Get-Date -Format dd-MMM-yyyy)"
}

# Reporting thresholds (percentages)
<#
    >= Good - reports as green in the progress bar
    >= Warning - reports as amber in the progress bar
    < Warning - reports as red in the progress bar
#>
$script:Thresholds = @{}
$Thresholds.Good = 90
$Thresholds.Warning = 70
$Thresholds.Inventory = @{} # Inventory thresholds are applicable to HW inventory, SW inventory and Heartbeat (DDR) only
$Thresholds.Inventory.Good = 90
$Thresholds.Inventory.Warning = 70

# Enable reports Regions
$active_Clients_Policy_Request = $True
$active_Clients_software_inventory = $True
$active_Clients_Hardware_Inventory = $True
$Client_Versions = $True
$System_Without_Client = $True
$Windows_Clients_Install_failures = $True
$Computer_Not_Rebooted = $True

#endregion

    #######################################################################
    #region Functions
    #######################################################################
    # Function to run a sql query
    function Get-SQLData 
    {
      param($Query)
      $connectionString = "Server=$dataSource;Database=$database;Integrated Security=SSPI;"
      $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
      $connection.ConnectionString = $connectionString
      $connection.Open()
    
      $command = $connection.CreateCommand()
      $command.CommandText = $Query
      $reader = $command.ExecuteReader()
      $table = New-Object -TypeName 'System.Data.DataTable'
      $table.Load($reader)
    
      # Close the connection
      $connection.Close()
    
      return $table
    }

    # Function to set the progress bar colour based on the the threshold value
    function Set-PercentageColour 
    {
      param(
        [int]$Value,
        [switch]$UseInventoryThresholds
      )

      If ($UseInventoryThresholds)
      {
        $Good = $Thresholds.Inventory.Good
        $Warning = $Thresholds.Inventory.Warning
      }
      Else
      {
        $Good = $Thresholds.Good
        $Warning = $Thresholds.Warning      
      }

      If ($Value -ge $Good)
      {
        $Hex = '#52B431' # Green
      }

      If ($Value -ge $Warning -and $Value -lt $Good)
      {
        $Hex = '#ff9900' # Amber
      }

      If ($Value -lt $Warning)
      {
        $Hex = '#FF0000' # Red
      }

      Return $Hex
    }
    #endregion

#######################################################################
#region Query Region
#######################################################################
# Create has table to store data
$Data = @{}

###########################################
# Discovered Systems With Clients Installed
###########################################

  $Query = "
    Select count(ResourceID) as 'Count' from v_R_System where (Client0 = 1)
  "
  $Data.ClientCount = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count
  
  # Get No Client Count
  $Query = "
    Select count(ResourceID) as 'Count' from v_R_System where (Client0 = 0 or Client0 is null) and Unknown0 is null
  "
  $Data.NoClientCount = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count
  
  # Calculate Client Percentage
  $Data.ClientCountPercentage = [Math]::Round($Data.ClientCount / ($Data.ClientCount + $Data.NoClientCount) * 100)
  $Data.NoClientCountPercentage = 100 - $Data.ClientCountPercentage
  $Data.TotalDiscoveredSystems = $Data.ClientCount + $Data.NoClientCount

###########################################
# Active Clients
###########################################

  $Query = "
    Select count(ResourceID) as 'Count' from v_CH_ClientSummary where ClientActiveStatus = 1
  "
  $Data.ActiveCount = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count
  
  # Get InActive Count
  $Query = "
    Select count(ResourceID) as 'Count' from v_CH_ClientSummary where ClientActiveStatus = 0
  "
  $Data.InactiveCount = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count
  
  # Calculate Active Percentage
  $Data.ActiveCountPercentage = [Math]::Round($Data.ActiveCount / ($Data.ActiveCount + $Data.InactiveCount) * 100)
  $Data.InActiveCountPercentage = 100 - $Data.ActiveCountPercentage
  $Data.ActiveInactiveTotal = $Data.ActiveCount + $Data.InactiveCount

###########################################
# Active/Pass Clients
###########################################

  $Query = "
    Select count(ResourceID) as 'Count' from v_CH_ClientSummary where ClientStateDescription = 'Active/Pass'
  "
  $Data.ActivePassCount = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count
  
  # Get Active/Fail Count
  $Query = "
    Select count(ResourceID) as 'Count' from v_CH_ClientSummary where ClientStateDescription = 'Active/Fail'
  "
  $Data.ActiveFailCount = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count
  
  # Get Active/Unknown Count
  $Query = "
    Select count(ResourceID) as 'Count' from v_CH_ClientSummary where ClientStateDescription = 'Active/Unknown'
  "
  $Data.ActiveUnknownCount = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count
  
  # Calculate Active/Pass Percentage
  $Data.ActivePassCountPercentage = [Math]::Round($Data.ActivePassCount / ($Data.ActivePassCount + $Data.ActiveFailCount + $Data.ActiveUnknownCount) * 100)
  $Data.ActiveNotPassCountPercentage = 100 - $Data.ActivePassCountPercentage

###########################################
# Active DDR
###########################################

  $Query = "
    Select count(ResourceID) as 'Count' from v_CH_ClientSummary where IsActiveDDR = 1 and ClientActiveStatus = 1
  "
  $Data.ActiveDDRCount = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count
  
  # Get InActive DDR Count
  $Query = "
    Select count(ResourceID) as 'Count' from v_CH_ClientSummary where IsActiveDDR = 0 and ClientActiveStatus = 1
  "
  $Data.InActiveDDRCount = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count
  
  # Calculate Active DDR Percentage
  $Data.ActiveDDRCountPercentage = [Math]::Round($Data.ActiveDDRCount / ($Data.ActiveDDRCount + $Data.InActiveDDRCount) * 100)
  $Data.InActiveDDRCountPercentage = 100 - $Data.ActiveDDRCountPercentage

############################################
# Hardware Inventory
############################################

  $Query = "
    Select count(ResourceID) as 'Count' from v_CH_ClientSummary where IsActiveHW = 1 and ClientActiveStatus = 1
  "
  $Data.ActiveHWCount = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count

  # Get InActive HW Count
  $Query = "
    Select count(ResourceID) as 'Count' from v_CH_ClientSummary where IsActiveHW = 0 and ClientActiveStatus = 1
  "
  $Data.InActiveHWCount = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count

  # Calculate Active HW Percentage
  $Data.ActiveHWCountPercentage = [Math]::Round($Data.ActiveHWCount / ($Data.ActiveHWCount + $Data.InActiveHWCount) * 100)
  $Data.InActiveHWCountPercentage = 100 - $Data.ActiveHWCountPercentage

############################################
# Software Inventory
############################################  
  
  $Query = "
    Select count(ResourceID) as 'Count' from v_CH_ClientSummary where IsActiveSW = 1 and ClientActiveStatus = 1
  "
  $Data.ActiveSWCount = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count

  # Get InActive SW Count
  $Query = "
    Select count(ResourceID) as 'Count' from v_CH_ClientSummary where IsActiveSW = 0 and ClientActiveStatus = 1
  "
  $Data.InActiveSWCount = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count

  # Calculate Active SW Percentage
  $Data.ActiveSWCountPercentage = [Math]::Round($Data.ActiveSWCount / ($Data.ActiveSWCount + $Data.InActiveSWCount) * 100)
  $Data.InActiveSWCountPercentage = 100 - $Data.ActiveSWCountPercentage

############################################
# Policy Request
############################################ 

  $Query = "
    Select count(ResourceID) as 'Count' from v_CH_ClientSummary where IsActivePolicyRequest = 1 and ClientActiveStatus = 1
  "
  $Data.ActivePRCount = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count

  # Get InActive PolicyRequest Count
  $Query = "
    Select count(ResourceID) as 'Count' from v_CH_ClientSummary where IsActivePolicyRequest = 0 and ClientActiveStatus = 1
  "
  $Data.InActivePRCount = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count

  # Calculate Active PolicyRequest Percentage
  $Data.ActivePRCountPercentage = [Math]::Round($Data.ActivePRCount / ($Data.ActivePRCount + $Data.InActivePRCount) * 100)
  $Data.InActivePRCountPercentage = 100 - $Data.ActivePRCountPercentage  

############################################
# Client Version
############################################ 

  $Query = "
    Select sys.Client_Version0 as 'Client Version', count (sys.ResourceID) as 'Count' from v_R_System sys
    inner join v_CH_ClientSummary ch on sys.ResourceID = ch.ResourceID
    where ch.ClientActiveStatus = 1
    Group by sys.Client_Version0
    Order by sys.Client_Version0 desc
  "
  $Data.ClientVersions = Get-SQLData -Query $Query
  $Data.TotalForClientVersions = [int]0
  $Data.ClientVersions | ForEach-Object -Process {
    $Data.TotalForClientVersions = $Data.TotalForClientVersions + $_.Count
  }
    #endregion



    # Create html header
    $html = @"
<!DOCTYPE html>
<html>
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="stylesheet" href="http://www.w3schools.com/lib/w3.css">
<body>

<style>
    body
  {
      background-color: Gainsboro;
  }

    table, th, td{
      border: 5px ;
      background-color: white;
       
      
    }

    h2,h3,h4{
        background-color:Darkblue;
        color:white;
        text-align: Left;
    }
    tr {
        cellpadding: 5px;
        }

     td {
        padding-top:2px;
        padding-bottom:2px
        padding:2px
        }
</style>


"@

    #endregion

    #######################################################################
    #region HTML - Discoverd Systems with client & Active Clients
    #######################################################################

    # Set html
    $html = $html + @"

<br>
<hl>
<br>
<table width="600" border="1" cellpadding="10" cellspacing="30">
  <tbody>
    <tr>
      <td><table cellpadding="1" cellspacing="1" width="400">
        <tr><h4> Discovered Systems with Client Installed</h4></tr>
        <tr>
          <td style="background-color:$(Set-PercentageColour -Value $Data.ClientCountPercentage);color:#ffffff;" width="$($Data.ClientCountPercentage)%"> $($Data.ClientCountPercentage)% </td>
          <td style="background-color:#eeeeee;color:#333333;" width="$($Data.NoClientCountPercentage)%"></td>
        </tr>
      </table>
        <table cellpadding="0" cellspacing="0" width="400">
          <tr>
            <td width="80%"> Discovered Systems with Client </td>
            <td width="20%"> $($Data.ClientCount) </td>
          </tr>
          <tr>
            <td width="80%"> Discovered Systems without Client </td>
            <td width="20%"> $($Data.NoClientCount) </td>
          </tr>
          <tr>
            <td width="80%"> Total </td>
            <td width="20%"> $($Data.TotalDiscoveredSystems) </td>
          </tr>
      </table></td>
      <td><h4>Active Clients</h4>
        <table cellpadding="0" cellspacing="0" width="400">
          <tr>
            <td style="background-color:$(Set-PercentageColour -Value $Data.ActiveCountPercentage);color:#ffffff;" width="$($Data.ActiveCountPercentage)%"> $($Data.ActiveCountPercentage)% </td>
            <td style="background-color:#eeeeee;color:#333333;" width="$($Data.InactiveCountPercentage)%"></td>
          </tr>
        </table>
        <table cellpadding="0" cellspacing="0" width="400">
          <tr>
            <td width="80%"> Active Clients </td>
            <td width="20%"> $($Data.ActiveCount) </td>
          </tr>
          <tr>
            <td width="80%"> Inactive Clients </td>
            <td width="20%"> $($Data.InActiveCount) </td>
          </tr>
          <tr>
            <td width="80%"> Total </td>
            <td width="20%"> $($Data.ActiveInactiveTotal) </td>
          </tr>
      </table></td>
    </tr>
  </tbody>
</table>
"@

    #endregion

    #######################################################################
    #region HTML Active Clients Health Evaluation & Active Clients Heartbeat (DDR)
    #######################################################################


    # Set html
    $html = $html + @"
<table width="600" border="1" cellpadding="10" cellspacing="30">
  <tbody>
    <tr>
      <td><h4>Active Clients Health Evaluation</h4>
        <table cellpadding="0" cellspacing="0" width="400">
            <tr>
                <td style="background-color:$(Set-PercentageColour -Value $Data.ActivePassCountPercentage);color:#ffffff;" width="$($Data.ActivePassCountPercentage)%">
                $($Data.ActivePassCountPercentage)%
                </td>
                <td style="background-color:#eeeeee;color:#333333;" width="$($Data.ActiveNotPassCountPercentage)%">
                </td>
            </tr>
        </table>
        <table cellpadding="0" cellspacing="0" width="400">
            <tr>
                <td width="80%">
                Active/Pass
                </td>
                <td width="20%">
                $($Data.ActivePassCount)
                </td>
            </tr>
            <tr>
                <td width="80%">
                Active/Fail
                </td>
                <td width="20%">
                $($Data.ActiveFailCount)
                </td>
            </tr>
            <tr>
                <td width="80%">
                Active/Unknown
                </td>
                <td width="20%">
                $($Data.ActiveUnknownCount)
                </td>
            </tr>
            </table></td>
      <td><h4>Active Clients Heartbeat (DDR)</h4>
        <table cellpadding="0" cellspacing="0" width="400">
          <tr>
            <td style="background-color:$(Set-PercentageColour -Value $Data.ActiveCountPercentage);color:#ffffff;" width="$($Data.ActiveCountPercentage)%"> $($Data.ActiveCountPercentage)% </td>
            <td style="background-color:#eeeeee;color:#333333;" width="$($Data.InactiveCountPercentage)%"></td>
          </tr>
        </table>
        <table cellpadding="0" cellspacing="0" width="400">
          <tr>
            <td width="80%"> Active Clients </td>
            <td width="20%"> $($Data.ActiveCount) </td>
          </tr>
          <tr>
            <td width="80%"> Inactive Clients </td>
            <td width="20%"> $($Data.InActiveCount) </td>
          </tr>
          <tr>
            <td width="80%"> Total </td>
            <td width="20%"> $($Data.ActiveInactiveTotal) </td>
          </tr>
      </table></td>
    </tr>
  </tbody>
</table>

"@
    #endregion

    #######################################################################
    #region Hardware Inventory & Software Inventory
    #######################################################################

    # Set html
    $html = $html + @"
<table width="600" border="1" cellpadding="10" cellspacing="30">
  <tbody>
    <tr>
        <td><h4>Active Clients Hardware Inventory</h4>
      <table cellpadding="0" cellspacing="0" width="400">
        <tr>
          <td style="background-color:$(Set-PercentageColour -Value $Data.ActiveHWCountPercentage -UseInventoryThresholds);color:#ffffff;" width="$($Data.ActiveHWCountPercentage)%">
          $($Data.ActiveHWCountPercentage)%
          </td>
          <td style="background-color:#eeeeee;color:#333333;" width="$($Data.InActiveHWCountPercentage)%">
          </td>
        </tr>
      </table>
      <table cellpadding="0" cellspacing="0" width="400">
        <tr>
            <td width="80%">
            Active HW Inventory
            </td>
            <td width="20%">
            $($Data.ActiveHWCount)
            </td>
        </tr>
        <tr>
            <td width="80%">
            Inactive HW Inventory
            </td>
            <td width="20%">
            $($Data.InActiveHWCount)
            </td>
        </tr>
      </table>
      </td>
      <td><h4>Active Clients Software Inventory</h4>
      <table cellpadding="0" cellspacing="0" width="400">
        <tr>
          <td style="background-color:$(Set-PercentageColour -Value $Data.ActiveSWCountPercentage -UseInventoryThresholds);color:#ffffff;" width="$($Data.ActiveSWCountPercentage)%">
          $($Data.ActiveSWCountPercentage)%
          </td>
          <td style="background-color:#eeeeee;color:#333333;" width="$($Data.InActiveSWCountPercentage)%">
          </td>
        </tr>
      </table>
      <table cellpadding="0" cellspacing="0" width="400">
        <tr>
            <td width="80%">
            Active SW Inventory
            </td>
            <td width="20%">
            $($Data.ActiveSWCount)
            </td>
        </tr>
        <tr>
            <td width="80%">
            Inactive SW Inventory
            </td>
            <td width="20%">
            $($Data.InActiveSWCount)
            </td>
        </tr>
      </table>
      </td>
    </tr>
  </tbody>
</table>
"@
    #endregion

    #######################################################################
    #region Get Active PolicyRequest Count & Client Version
    #######################################################################

    # Set html
    $html = $html + @"
<table width="930" border="1" cellspacing="30" cellpadding="10" bordercolor="black">
  <tbody>
    <tr>
      <td width="400">
      <table cellpadding="0" cellspacing="0">
        <tr><h4>Active Clients Policy Request</h4></tr>
      </table>
          <table cellpadding="0" cellspacing="0" width="400">
            <tr>
              <td style="background-color:$(Set-PercentageColour -Value $Data.ActivePRCountPercentage);color:#ffffff;" width="$($Data.ActivePRCountPercentage)%">
              $($Data.ActivePRCountPercentage)%
              </td>
              <td style="background-color:#eeeeee;color:#333333;" width="$($Data.InActivePRCountPercentage)%">
              </td>
            </tr>
          </table>
          <table cellpadding="0" cellspacing="0" width="400">
            <tr>
                <td width="80%">
                Active Policy Request
                </td>
                <td width="20%">
                $($Data.ActivePRCount)
                </td>
            </tr>
            <tr>
                <td width="80%">
                Inactive Policy Request
                </td>
                <td width="20%">
                $($Data.InActivePRCount)
                </td>
            </tr>
          </table>
        </td>
<td>
    
      </td>
    </tr>
  </tbody>
</table>
"@



    #endregion

    #######################################################################
    #region Get Client Versions
    #######################################################################



# Set html
$html = $html + @"
<br><br>
<h2>Client Versions</h2>
<table cellpadding="0" cellspacing="0" width="400">
<tr>
    <th style="text-align: left" width="240">Version</th>
    <th style="text-align: left" width="80">Count</th>
    <th style="text-align: left" width="80">Percent</th>
</tr>
</table>
"@

$Data.ClientVersions | foreach {
$Percentage = [Math]::Round($_.Count / $Data.TotalForClientVersions * 100)
$PercentageRemaining = (100 - $Percentage)
$html = $html + @"
<table cellpadding="0" cellspacing="0" width="400">
<tr>
    <td width="240">
    $($_.'Client Version')
    </td>
    <td width="80">
    $($_.Count)
    </td>
    <td width="80">
    $($Percentage)%
    </td>
</tr>
</table>
"@
}
#endregion

#######################################################################
#region Get No Client Systems
#######################################################################

if ($System_Without_Client -eq $True)

{
  $Data.NoClient = @{}
  # no client - unknown OS
  $Query = "
    Select count(ResourceID) as 'Count' from v_R_System 
    where (Client0 = 0 or Client0 is null) 
    and Unknown0 is null
    and Operating_System_Name_and0 like 'unknown%'
  "
  $Data.NoClient.UnknownOS = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count

  # no client windows OS
  $Query = "
    Select count(ResourceID) as 'Count' from v_R_System 
    where (Client0 = 0 or Client0 is null) 
    and Unknown0 is null
    and Operating_System_Name_and0 like '%Windows%'
  "
  $Data.NoClient.WindowsOS = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count

  # no client other OS
  $Query = "
    Select count(ResourceID) as 'Count' from v_R_System 
    where (Client0 = 0 or Client0 is null) 
    and Unknown0 is null
    and Operating_System_Name_and0 not like '%Windows%'
    and Operating_System_Name_and0 not like 'unknown%'
  "
  $Data.NoClient.OtherOS = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count

  # no client and no last logon timestamp in last 7 days
  $Query = "
    Select count(ResourceID) as 'Count' from v_R_System 
    where (Client0 = 0 or Client0 is null) 
    and Unknown0 is null
    and (DATEDIFF(day,Last_Logon_Timestamp0, GetDate())) >= 7
  "
  $Data.NoClient.GTLast7 = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count

  # no client and last logon timestamp within last 7 days
  $Query = "
    Select count(ResourceID) as 'Count' from v_R_System 
    where (Client0 = 0 or Client0 is null) 
    and Unknown0 is null
    and (DATEDIFF(day,Last_Logon_Timestamp0, GetDate())) < 7
  "
  $Data.NoClient.LTLast7 = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count

  # Set html
  $html = $html + @"
<br><br>
<h3>Systems with No Client</h2>
<table cellpadding="0" cellspacing="0" width="400">
<tr>
    <th style="text-align: left" width="240">Category</th>
    <th style="text-align: left" width="80">Count</th>
    <th style="text-align: left" width="80">Percent</th>
</tr>
</table>
"@

  $html = $html + @"
<table cellpadding="0" cellspacing="0" width="400">
<tr>
    <td width="240">
    Windows OS
    </td>
    <td width="80">
    $($Data.NoClient.WindowsOS)
    </td>
    <td width="80">
    $([Math]::Round($Data.NoClient.WindowsOS / $Data.NoClientCount * 100))%
    </td>
</tr>
<tr>
    <td width="240">
    Other OS
    </td>
    <td width="80">
    $($Data.NoClient.OtherOS)
    </td>
    <td width="80">
    $([Math]::Round($Data.NoClient.OtherOS / $Data.NoClientCount * 100))%
    </td>
</tr>
<tr>
    <td width="240">
    Unknown OS
    </td>
    <td width="80">
    $($Data.NoClient.UnknownOS)
    </td>
    <td width="80">
    $([Math]::Round($Data.NoClient.UnknownOS / $Data.NoClientCount * 100))%
    </td>
</tr>
<tr>
    <td width="240">
    Last Logon > 7 days
    </td>
    <td width="80">
    $($Data.NoClient.GTLast7)
    </td>
    <td width="80">
    $([Math]::Round($Data.NoClient.GTLast7 / $Data.NoClientCount * 100))%
    </td>
</tr>
<tr>
    <td width="240">
    Last Logon < 7 days
    </td>
    <td width="80">
    $($Data.NoClient.LTLast7)
    </td>
    <td width="80">
    $([Math]::Round($Data.NoClient.LTLast7 / $Data.NoClientCount * 100))%
    </td>
</tr>
</table>
"@
}
#endregion

#######################################################################
#region Windows Client Installation Failures
#######################################################################

if ($Windows_Clients_Install_failures -eq $True)

{
  $Query = "
    select count(cdr.MachineID) as 'Count',
    cdr.CP_LastInstallationError as 'Error Code'
    from v_CombinedDeviceResources cdr
    where 
    cdr.IsClient = 0
    and cdr.DeviceOS like '%Windows%'
    group by cdr.CP_LastInstallationError 
  "
  $InstallErrors = Get-SQLData -Query $Query

  # Translate error codes to friendly names and add the percentage value
  $TotalErrors = 0
  $InstallErrors | ForEach-Object -Process {
    $TotalErrors = $TotalErrors + $_.Count
  }
  $Data.InstallFailures = $InstallErrors |
  Select-Object -Property Count, 'Error Code', @{
    n = 'Error Description'
    e = {
      ([ComponentModel.Win32Exception]$_.'Error Code').Message
    }
  }, @{
    n = 'Percentage'
    e = {
      [Math]::Round($_.Count / $TotalErrors * 100)
    }
  } |
  Sort-Object -Property Count -Descending

  # Set html
  $html = $html + @"
<br><br>
<h3>Windows Client Installation Failures</h2>
<table cellpadding="0" cellspacing="0" width="600">
<tr>
    <th style="text-align: left" width="100">Error Code</th>
    <th style="text-align: left" width="300">Error Description</th>
    <th style="text-align: left" width="100">Count</th>
    <th style="text-align: left" width="100">Percent</th>
</tr>
</table>
"@

  $Data.InstallFailures | ForEach-Object -Process {
    $html = $html + @"
<table cellpadding="0" cellspacing="0" width="600">
<tr>
    <td width="100">
    $($_.'Error Code')
    </td>
    <td width="300">
    $($_.'Error Description')
    </td>
    <td width="100">
    $($_.Count)
    </td>
    <td width="100">
    $($_.Percentage)%
    </td>
</tr>
</table>
"@
  }
}
#endregion

#######################################################################
#region Computer reboots
#######################################################################

if ($Computer_Not_Rebooted -eq $True)

{
  # Active systems with a known Last BootUp Time
  $Query = "
    select Count(ch.ResourceID) as 'Count'
    from v_CH_ClientSummary ch
    left join v_GS_OPERATING_SYSTEM os on os.ResourceId = ch.ResourceId 
    where os.LastBootUpTime0 is not null
    and ch.ClientActiveStatus = 1
  "
  $Data.ActiveLastBootUpTotal = Get-SQLData -Query $Query | Select-Object -ExpandProperty Count

  # Computer reboot dates
  $Query = "
    select '7 days' as TimePeriod,Count(sys.Name0) as 'Count',1 SortOrder
    from v_R_System sys
    inner join v_GS_OPERATING_SYSTEM os on os.ResourceId = sys.ResourceId 
    inner join v_CH_ClientSummary ch on ch.ResourceID = sys.ResourceID
    where os.LastBootUpTime0 < DATEADD(day,-7, GETDATE())
    and ch.ClientActiveStatus = 1
    UNION
    select '14 days' as TimePeriod,Count(sys.Name0) as 'Count',2
    from v_R_System sys
    inner join v_GS_OPERATING_SYSTEM os on os.ResourceId = sys.ResourceId 
    inner join v_CH_ClientSummary ch on ch.ResourceID = sys.ResourceID
    where os.LastBootUpTime0 < DATEADD(day,-14, GETDATE())
    and ch.ClientActiveStatus = 1
    UNION
    select '1 month' as TimePeriod,Count(sys.Name0) as 'Count',3
    from v_R_System sys
    inner join v_GS_OPERATING_SYSTEM os on os.ResourceId = sys.ResourceId 
    inner join v_CH_ClientSummary ch on ch.ResourceID = sys.ResourceID
    where os.LastBootUpTime0 < DATEADD(month,-1, GETDATE())
    and ch.ClientActiveStatus = 1
    UNION
    select '3 months' as TimePeriod,Count(sys.Name0) as 'Count',4
    from v_R_System sys
    inner join v_GS_OPERATING_SYSTEM os on os.ResourceId = sys.ResourceId 
    inner join v_CH_ClientSummary ch on ch.ResourceID = sys.ResourceID
    where os.LastBootUpTime0 < DATEADD(MONTH,-3, GETDATE())
    and ch.ClientActiveStatus = 1
    UNION
    select '6 months' as TimePeriod,Count(sys.Name0) as 'Count',5
    from v_R_System sys
    inner join v_GS_OPERATING_SYSTEM os on os.ResourceId = sys.ResourceId 
    inner join v_CH_ClientSummary ch on ch.ResourceID = sys.ResourceID
    where os.LastBootUpTime0 < DATEADD(MONTH,-6, GETDATE())
    and ch.ClientActiveStatus = 1
    Order By SortOrder
  "
  $Data.ComputerReboots = Get-SQLData -Query $Query

  # Set html
  $html = $html + @"
<br><br>
<h3>Computers Not Rebooted</h2>
<table cellpadding="0" cellspacing="0" width="400">
<tr>
    <th style="text-align: left" width="240">Time Period</th>
    <th style="text-align: left" width="80">Count</th>
    <th style="text-align: left" width="80">Percent*</th>
</tr>
</table>
"@

  $html = $html + @"
<table cellpadding="0" cellspacing="0" width="400">
<tr>
    <td width="240">
    $($Data.ComputerReboots[0].TimePeriod)
    </td>
    <td width="80">
    $($Data.ComputerReboots[0].Count)
    </td>
    <td width="80">
    $([Math]::Round($Data.ComputerReboots[0].Count / $Data.ActiveLastBootUpTotal * 100))%
    </td>
</tr>
<tr>
    <td width="240">
    $($Data.ComputerReboots[1].TimePeriod)
    </td>
    <td width="80">
    $($Data.ComputerReboots[1].Count)
    </td>
    <td width="80">
    $([Math]::Round($Data.ComputerReboots[1].Count / $Data.ActiveLastBootUpTotal * 100))%
    </td>
</tr>
<tr>
    <td width="240">
    $($Data.ComputerReboots[2].TimePeriod)
    </td>
    <td width="80">
    $($Data.ComputerReboots[2].Count)
    </td>
    <td width="80">
    $([Math]::Round($Data.ComputerReboots[2].Count / $Data.ActiveLastBootUpTotal * 100))%
    </td>
</tr>
<tr>
    <td width="240">
    $($Data.ComputerReboots[3].TimePeriod)
    </td>
    <td width="80">
    $($Data.ComputerReboots[3].Count)
    </td>
    <td width="80">
    $([Math]::Round($Data.ComputerReboots[3].Count / $Data.ActiveLastBootUpTotal * 100))%
    </td>
</tr>
<tr>
    <td width="240">
    $($Data.ComputerReboots[4].TimePeriod)
    </td>
    <td width="80">
    $($Data.ComputerReboots[4].Count)
    </td>
    <td width="80">
    $([Math]::Round($Data.ComputerReboots[4].Count / $Data.ActiveLastBootUpTotal * 100))%
    </td>
</tr>
</table>
<div style="font-size: 12px">* Percentage is calculated from the total number of active clients that have a known last bootup time ($($Data.ActiveLastBootUpTotal))</div>
"@
}
#endregion

#######################################################################
#region Client Health Thresholds
#######################################################################

$Query = '
  SELECT *
  FROM v_CH_Settings
  where SettingsID = 1
'

$Data.CHSettings = Get-SQLData -Query $Query

# Set html
$html = $html + @"
<br><br>
<h3>Client Status Settings</h2>
<table cellpadding="0" cellspacing="0" width="400">
<tr>
    <th style="text-align: left" width="300">Setting</th>
    <th style="text-align: left" width="100">Days</th>
</tr>
</table>
"@

$html = $html + @"
<table cellpadding="0" cellspacing="0" width="400">
<tr>
    <td width="300">
    Heartbeat Discovery
    </td>
    <td width="100">
    $($Data.CHSettings.DDRInactiveInterval)
    </td>
</tr>
<tr>
    <td width="300">
    Hardware Inventory
    </td>
    <td width="100">
    $($Data.CHSettings.HWInactiveInterval)
    </td>
</tr>
<tr>
    <td width="300">
    Software Inventory
    </td>
    <td width="100">
    $($Data.CHSettings.SWInactiveInterval)
    </td>
</tr>
<tr>
    <td width="300">
    Policy Requests
    </td>
    <td width="100">
    $($Data.CHSettings.PolicyInactiveInterval)
    </td>
</tr>
<tr>
    <td width="300">
    Status History Retention
    </td>
    <td width="100">
    $($Data.CHSettings.CleanUpInterval)
    </td>
</tr>

</table>
"@
#endregion

#######################################################################
#region Maintenance Task status
#######################################################################

$Query = '
  select *,
  floor(DATEDIFF(ss,laststarttime,lastcompletiontime)/3600) as Hours,
  floor(DATEDIFF(ss,laststarttime,lastcompletiontime)/60)- floor(DATEDIFF(ss,laststarttime,lastcompletiontime)/3600)*60 as Minutes,
  floor(DATEDIFF(ss,laststarttime,lastcompletiontime))- floor(DATEDIFF(ss,laststarttime,lastcompletiontime)/60)*60 as TotalSeconds
  from SQLTaskStatus
'

$Data.MWStatus = Get-SQLData -Query $Query

# Set html
$html = $html + @"
<br><br>
<h3>Maintenance Task Status</h2>
<table cellpadding="0" cellspacing="0" width="1200">
<tr>
    <th style="text-align: left" width="500">Taskname</th>
    <th style="text-align: left" width="200">LastStartTime</th>
    <th style="text-align: left" width="200">LastCompletionTime</th>
    <th style="text-align: left" width="200">CompletionStatus</th>
</tr>
</table>
"@

$Data.MWStatus | ForEach-Object -Process {
  $html = $html + @"
<table cellpadding="0" cellspacing="0" width="1200">
<tr>
    <td style="text-align: left" width="500">
    $($_.'Taskname')
    </td>
         <td style="text-align: left" width="200">
    $($_.'LastStartTime')
    </td>
            <td style="text-align: left" width="200">
    $($_.'LastCompletionTime')
    </td>
                <td style="text-align: left" width="200">
    $($_.'Completionstatus')
    </td>
</tr>
</table>
"@
}
#endregion

#######################################################################
#region Database file size
#######################################################################

$Query = "
  select
  Sys.FILEID as 'FileID',
  left(Sys.NAME,15) as 'DBName',
  left(Sys.FILENAME,60) as 'DBFilePath',
  convert(decimal(12,2),round(Sys.size/128.000,2)) as 'Filesize_MB',
  convert(decimal(12,2),round(fileproperty(Sys.name,'SpaceUsed')/128.000,2)) as 'UsedSpace_MB',
  convert(decimal(12,2),round((Sys.size-fileproperty(Sys.name,'SpaceUsed'))/128.000,2)) as 'FreeSpace_MB',
  convert(decimal(12,2),round(Sys.growth/128.000,2)) as 'GrowthSpace_MB'
  from dbo.sysfiles Sys
"

$Data.DBStatus = Get-SQLData -Query $Query

# Set html
$html = $html + @"
<br><br>
<h3>Databasefiles Status</h2>
<table cellpadding="0" cellspacing="0" width="1200">
<tr>
    <th style="text-align: left" width="500">FileName</th>
    <th style="text-align: left" width="200">FileSize (MB)</th>
    <th style="text-align: left" width="200">UsedSpace (MB)</th>
    <th style="text-align: left" width="200">FreeSpace (MB)</th>
    <th style="text-align: left" width="200">GrowthSpace (MB)</th>
</tr>
</table>
"@

$Data.DBStatus | ForEach-Object -Process {
  $html = $html + (@'
<table cellpadding="0" cellspacing="0" width="1200">
<tr>
    <td style="text-align: Left" width="250">
    {0}
    </td>
         <td style="text-align: Right" width="180">
    {1}
    </td>
            <td style="text-align: Right" width="180">
    {2}
    </td>
                <td style="text-align: Right" width="180">
    {3}
    </td>
    <td style="text-align: Right" width="180">
    {4}
    </td>
</tr>
</table>
'@ -f $_.'DBName', $_.'FileSize_MB', $_.'UsedSpace_MB', $_.'FreeSpace_MB', $_.'GrowthSpace_MB')
}
#endregion

#######################################################################
#region Close html document...
#######################################################################
$html = $html + @"
</body>
</html>
"@
#endregion

#######################################################################
# Send email
#######################################################################

#Send-MailMessage @EmailParams -Body $html -BodyAsHtml

#######################################################################
# Test, enable this row to generate html-page
#######################################################################

$html | Out-File -FilePath C:\Temp\test2.html
