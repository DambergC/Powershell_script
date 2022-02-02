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

# Siteconfiguration
$sitecode = 'ps1'
$siteserver = 'TH-mgt02.korsberga.local'
$StatusMessageTime = (Get-Date).AddDays(-7)

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
#region Powershell Region
#######################################################################

#######################################
# Powershell - ADR Status
#######################################
$ADRstatus = Get-CMSoftwareUpdateAutoDeploymentRule -Fast
  
#######################################
# Powershell - StatusMessage Siteserver
#######################################
$StatusMessageSiteserver = Get-CMSiteStatusMessage -ComputerName $siteserver -Severity Error -SiteCode $sitecode -StartDateTime $StatusMessageTime
  
#######################################
# Powershell - Disk status Siteserver
#######################################
$DiskReport = Get-CimInstance win32_logicaldisk -Filter "Drivetype=3" -ErrorAction SilentlyContinue | Select-Object `
@{Label = "HostName"; Expression = { $_.SystemName } },
@{Label = "DriveLetter"; Expression = { $_.DeviceID } },
@{Label = "DriveName"; Expression = { $_.VolumeName } },
@{Label = "Total Capacity (GB)"; Expression = { "{0:N1}" -f ( $_.Size / 1gb) } },
@{Label = "Free Space (GB)"; Expression = { "{0:N1}" -f ( $_.Freespace / 1gb ) } },
@{Label = 'Free Space (%)'; Expression = { "{0:P0}" -f ($_.Freespace / $_.Size) } } 

  
#endregion


#######################################################################
#region Query Region
#######################################################################
# Create has table to store data
$Data = @{}

###########################################
# QUERY - Discovered Systems With Clients Installed
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
# QUERY - Active Clients
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
# QUERY - Active/Pass Clients
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
# QUERY - Active DDR
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
# QUERY - Hardware Inventory
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
# QUERY - Software Inventory
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
# QUERY - Policy Request
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
# QUERY - Client Version
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


############################################
# QUERY - No Client System
############################################

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

############################################
# QUERY - Client Installation failure
############################################

$Query = "
  select count(cdr.MachineID) as 'Count',
  cdr.CP_LastInstallationError as 'Error Code'
  from v_CombinedDeviceResources cdr
  where 
  cdr.IsClient = 0
  and cdr.DeviceOS like '%Windows%'
  and CP_LastInstallationError >= 0
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

############################################
# QUERY - Computer Reboots
############################################

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

############################################
# QUERY - Client Health Thresholds
############################################

$Query = '
  SELECT *
  FROM v_CH_Settings
  where SettingsID = 1
'

$Data.CHSettings = Get-SQLData -Query $Query 

############################################
# QUERY - Maintenance Task Status
############################################

$Query = '
  select *,
  floor(DATEDIFF(ss,laststarttime,lastcompletiontime)/3600) as Hours,
  floor(DATEDIFF(ss,laststarttime,lastcompletiontime)/60)- floor(DATEDIFF(ss,laststarttime,lastcompletiontime)/3600)*60 as Minutes,
  floor(DATEDIFF(ss,laststarttime,lastcompletiontime))- floor(DATEDIFF(ss,laststarttime,lastcompletiontime)/60)*60 as TotalSeconds
  from SQLTaskStatus
'

$Data.MWStatus = Get-SQLData -Query $Query

############################################
# QUERY - Database File size
############################################

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

#endregion

#######################################################################
#region Create html header
#######################################################################

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
      align: Left;
      text-align: left;
      padding-left: 10px;
      padding-right: 10px;
       
      
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
        text-align: left;
        }
</style>


"@

#endregion

#######################################################################
#region HTML ADR Status
#######################################################################

# Set html
$html = $html + @"
    <table width="930" border="1" cellspacing="1">
    <tbody>
        <tr>
            <td>
            <h4>Automatic Deployment Rules - Status</h4>
        
        <table width="100%">
        <tr>
            <td width="2%"></td>
            <th width="30%" >ADR Name</th>
            <th width="20%" >Last Error Code</th>
            <th width="25%" >Last Error Time</th>
            <th width="25%" >Last Run Time</th>
        </tr>
        </table>
                    </td>
        </tr>
    </tbody>
    </table>
"@

$ADRstatus | ForEach-Object -Process {
  

  
  $html = $html + @"
    <table width="930" border="1" cellspacing="1">
    <tbody>
        <tr>
            <td>
        <table width="100%">      
        <tr  bgcolor="red">
            <td width="2%"></td>
            <td width="30%">
            $($_.name)
            </td>
            <td width="20%">
            $($_.lasterrorcode)
            </td>
            <td width="25%">
            $($_.lasterrortime)
            </td>
            <td width="25%">
            $($_.lastRuntime)
            </td>
        </tr>
                </table>
            </td>
        </tr>
    </tbody>
    </table>
"@
  

}


#endregion

#######################################################################
#region HTML StatusMessage Siteserver
#######################################################################

# Set html
$html = $html + @"
    <table width="930" border="1" cellspacing="1">
    <tbody>
        <tr>
            <td>
            <h4>StatusMessage Siteserver $siteserver - last 7 days</h4>
        
        <table width="100%">
        <tr>
            <td width="2%"></td>
            <th width="50%" >Component</th>
            <th width="20%" >MessageID</th>
            <th width="30%" >Time</th>

        </tr>
        </table>
                    </td>
        </tr>
    </tbody>
    </table>
"@

$StatusMessageSiteserver | ForEach-Object -Process {
  

  
  $html = $html + @"
    <table width="930" border="1" cellspacing="1">
    <tbody>
        <tr>
            <td>
        <table width="100%">      
        <tr  bgcolor="red">
            <td width="2%"></td>
            <td width="50%">
            $($_.component)
            </td>
            <td width="20%">
            $($_.Messageid)
            </td>
            <td width="30%">
            $($_.time)
            </td>

        </tr>
                </table>
            </td>
        </tr>
    </tbody>
    </table>
"@
  

}


#endregion

#######################################################################
#region HTML Diskspace Siteserver
#######################################################################

# Set html
$html = $html + @"
    <table width="930" border="1" cellspacing="1">
    <tbody>
        <tr>
            <td>
            <h4>Diskstatus $siteserver</h4>
        
        <table width="100%">
        <tr>
            <td width="2%"></td>
            <th width="20%" >DriveLetter</th>
            <th width="20%" >Description</th>
            <th width="20%" >Total Capacity (GB)</th>
            <th width="20%" >Free Space (GB)</th>
            <th width="20%" >Free Space (%)</th>
        </tr>
        </table>
                    </td>
        </tr>
    </tbody>
    </table>
"@

$DiskReport | ForEach-Object -Process {
  
  $html = $html + @"
    <table width="930" border="1" cellspacing="1">
    <tbody>
        <tr>
            <td>
        <table width="100%">      
        <tr  bgcolor="red">
            <td width="2%"></td>
            <td width="20%">
            $($_.driveletter)
            </td>
            <td width="20%">
            $($_.Drivename)
            </td>
            <td width="20%">
            $($_.'Total Capacity (GB)')
            </td>
            <td width="20%">
            $($_.'Free Space (GB)')
            </td>
            <td width="20%">
            $($_.'Free Space (%)')
            </td>

        </tr>
                </table>
            </td>
        </tr>
    </tbody>
    </table>
"@
  

}


#endregion


#######################################################################
#region HTML Client Versions
#######################################################################
    
# Set html
$html = $html + @"
<table width="930" border="1">
    <tbody>
	<tr>
	    <th><h4>Client Versions</h4>
        <table  width="100%">
        <tr>
            <th  width="60%">Version</th>
            <th  width="20%">Count</th>
            <th  width="20%">Percent</th>
        </tr>
        </table>
        </th>
    </tr>
    </tbody>
</table>
"@

$Data.ClientVersions | ForEach-Object -Process {
  $Percentage = [Math]::Round($_.Count / $Data.TotalForClientVersions * 100)
  $PercentageRemaining = (100 - $Percentage)
  $html = $html + @"
<table width="930">
    <tbody>
    <tr>
        <td width="2%"></td>
        <td width="60%">
        $($_.'Client Version')
        </td>
        <td width="20%">
        $($_.Count)
        </td>
        <td width="20%">
        $($Percentage)%
        </td>
    </tr>
    </tbody>
</table>
"@
}
#endregion

#######################################################################
#region HTML Get No Client Systems
#######################################################################

# Set html
$html = $html + @"
    <table width="930" border="1">
    <tbody>
        <tr>
            <td>
            <h4>Systems with No Client</h4>
                <table width="100%">
                <tr>
                    <th  width="60%">Category</th>
                    <th  width="20%">Count</th>
                    <th  width="20%">Percent</th>
                </tr>
                </table>
            </td>
        </tr>
    </tbody>
    </table>
"@

$html = $html + @"
    <table width="930" border="1">
    <tbody>
        <tr>
            <td>
                <table width="100%">
                <tr>
                    <td width="60%">
                    Windows OS
                    </td>
                    <td width="20%">
                    $($Data.NoClient.WindowsOS)
                    </td>
                    <td width="20%">
                    $([Math]::Round($Data.NoClient.WindowsOS / $Data.NoClientCount * 100))%
                    </td>
                </tr>
                <tr>
                    <td width="60%">
                    Other OS
                    </td>
                    <td width="20%">
                    $($Data.NoClient.OtherOS)
                    </td>
                    <td width="20%">
                    $([Math]::Round($Data.NoClient.OtherOS / $Data.NoClientCount * 100))%
                    </td>
                </tr>
                <tr>
                    <td width="60%">
                    Unknown OS
                    </td>
                    <td width="20%">
                    $($Data.NoClient.UnknownOS)
                    </td>
                    <td width="20%">
                    $([Math]::Round($Data.NoClient.UnknownOS / $Data.NoClientCount * 100))%
                    </td>
                </tr>
                <tr>
                    <td width="60%">
                    Last Logon > 7 days
                    </td>
                    <td width="20%">
                    $($Data.NoClient.GTLast7)
                    </td>
                    <td width="20%">
                    $([Math]::Round($Data.NoClient.GTLast7 / $Data.NoClientCount * 100))%
                    </td>
                </tr>
                <tr>
                    <td width="60%">
                    Last Logon < 7 days
                    </td>
                    <td width="20%">
                    $($Data.NoClient.LTLast7)
                    </td>
                    <td width="20%">
                    $([Math]::Round($Data.NoClient.LTLast7 / $Data.NoClientCount * 100))%
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
#region HTML Windows Client Installation Failures
#######################################################################


# Set html
$html = $html + @"
    <table width="930" border="1">
    <tbody>
        <tr>
            <td>
                <h4>Windows Client Installation Failures</h4>
                <table width="100%">
                <tr>
                    <th  width="20%">Error Code</th>
                    <th  width="60%">Error Description</th>
                    <th  width="10%">Count</th>
                    <th  width="10%">Percent</th>
                </tr>
                </table>
            </td>
        </tr>
    </tbody>
    </table>
"@

$Data.InstallFailures | ForEach-Object -Process {
  $html = $html + @"
    <table width="930" border="1">
    <tbody>
        <tr>
            <td>
                <table width="100%">
                <tr>
                    <td width="20%">
                    $($_.'Error Code')
                    </td>
                    <td width="60%">
                    $($_.'Error Description')
                    </td>
                    <td width="10%">
                    $($_.Count)
                    </td>
                    <td width="10%">
                    $($_.Percentage)%
                    </td>
                </tr>
                </table>
            </td>
        </tr>
    </tbody>
    </table>
"@
}

#endregion

#######################################################################
#region HTML Computer reboots
#######################################################################

# Set html
$html = $html + @"
    <table width="930" border="1">
    <tbody>
        <tr>
            <td>
                <h4>Computers Not Rebooted</h4>
                    <table width="100%">
                    <tr>
                        <th  width="60%">Time Period</th>
                        <th  width="20%">Count</th>
                        <th  width="20%">Percent*</th>
                    </tr>
                    </table>
"@

$html = $html + @"
                    <table width="100%">
                        <tr>
                            <td width="60%">
                            $($Data.ComputerReboots[0].TimePeriod)
                            </td>
                            <td width="20%">
                            $($Data.ComputerReboots[0].Count)
                            </td>
                            <td width="20%">
                            $([Math]::Round($Data.ComputerReboots[0].Count / $Data.ActiveLastBootUpTotal * 100))%
                            </td>
                        </tr>
                        <tr>
                            <td width="60%">
                            $($Data.ComputerReboots[1].TimePeriod)
                            </td>
                            <td width="20%">
                            $($Data.ComputerReboots[1].Count)
                            </td>
                            <td width="20%">
                            $([Math]::Round($Data.ComputerReboots[1].Count / $Data.ActiveLastBootUpTotal * 100))%
                            </td>
                        </tr>
                        <tr>
                            <td width="60%">
                            $($Data.ComputerReboots[2].TimePeriod)
                            </td>
                            <td width="20%">
                            $($Data.ComputerReboots[2].Count)
                            </td>
                            <td width="20%">
                            $([Math]::Round($Data.ComputerReboots[2].Count / $Data.ActiveLastBootUpTotal * 100))%
                            </td>
                        </tr>
                        <tr>
                            <td width="60%">
                            $($Data.ComputerReboots[3].TimePeriod)
                            </td>
                            <td width="20%">
                            $($Data.ComputerReboots[3].Count)
                            </td>
                            <td width="20%">
                            $([Math]::Round($Data.ComputerReboots[3].Count / $Data.ActiveLastBootUpTotal * 100))%
                            </td>
                        </tr>
                        <tr>
                            <td width="60%">
                            $($Data.ComputerReboots[4].TimePeriod)
                            </td>
                            <td width="20%">
                            $($Data.ComputerReboots[4].Count)
                            </td>
                            <td width="20%">
                            $([Math]::Round($Data.ComputerReboots[4].Count / $Data.ActiveLastBootUpTotal * 100))%
                            </td>
                        </tr>
                    </table>
                    <div style="font-size: 12px">* Percentage is calculated from the total number of active clients that have a known last bootup time ($($Data.ActiveLastBootUpTotal))</div>
                </td>
            </tr>
    </tbody>
    </table>
"@

#endregion

#######################################################################
#region HTML Client Health Thresholds
#######################################################################

# Set html
$html = $html + @"
    <table width="930" border="1">
    <tbody>
        <tr>
            <td>
                <h4>Client Status Settings</h4>
                    <table width="100%">
                    <tr>
                        <th  width="75%">Setting</th>
                        <th  width="25%">Days</th>
                    </tr>
                    </table>
"@

$html = $html + @"
                    <table width="100%">
                    <tr>
                        <td width="75%">
                        Heartbeat Discovery
                        </td>
                        <td width="25%">
                        $($Data.CHSettings.DDRInactiveInterval)
                        </td>
                    </tr>
                    <tr>
                        <td width="75%">
                        Hardware Inventory
                        </td>
                        <td width="25%">
                        $($Data.CHSettings.HWInactiveInterval)
                        </td>
                    </tr>
                    <tr>
                        <td width="75%">
                        Software Inventory
                        </td>
                        <td width="25%">
                        $($Data.CHSettings.SWInactiveInterval)
                        </td>
                    </tr>
                    <tr>
                        <td width="75%">
                        Policy Requests
                        </td>
                        <td width="25%">
                        $($Data.CHSettings.PolicyInactiveInterval)
                        </td>
                    </tr>
                    <tr>
                        <td width="75%">
                        Status History Retention
                        </td>
                        <td width="25%">
                        $($Data.CHSettings.CleanUpInterval)
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
#region HTML Maintenance Task status
#######################################################################

# Set html
$html = $html + @"
    <table width="930" border="1">
    <tbody>
    <tr>
        <td>
            <h4>Maintenance Task Status</h4>
            <table width="100%">
                <tr>
                    <th width="50%">Taskname</th>
                    <th width="25%">LastStartTime</th>
                    <th width="25%">LastCompletionTime</th>
                </tr>
            </table>
       </td>
    </tr>
    </tbody>
    </table>
"@

$Data.MWStatus | ForEach-Object -Process {
  $html = $html + @"
    <table width="930" border="1">
    <tbody>
    <tr>
        <td>
            <table width="100%">
                <tr>

                    <td width="50%">
                    $($_.'Taskname')
                    </td>
                         <td width="25%">
                    $($_.'LastStartTime')
                    </td>
                            <td width="25%">
                    $($_.'LastCompletionTime')
                    </td>

                </tr>
            </table>
        </td>
    </tr>
    </tbody>
    </table>
"@
}
#endregion

#######################################################################
#region HTML Database file size
#######################################################################

# Set html
$html = $html + @"
    <table width="930" border="1">
    <tbody>
    <tr>
        <td>
            <h4>Databasefiles Status</h4>
                <table width="100%">
                <tr>
                    <th width="40%">FileName</th>
                    <th width="15%">FileSize (MB)</th>
                    <th width="15%">UsedSpace (MB)</th>
                    <th width="15%">FreeSpace (MB)</th>
                    <th width="15%">GrowthSpace (MB)</th>
                </tr>
                </table>
        </td>
    </tr>
    </tbody>
    </table>

"@

$Data.DBStatus | ForEach-Object -Process {
  $html = $html + @"
  <table width="930" border="1">
                <tbody>
                <tr>
                    <td>
                        <table width="100%">
                        <tr>
                            <td width="40%">
                            $($_.'DBName')
                            </td>
                            <td width="15%">
                            $($_.'FileSize_MB')
                            </td>
                            <td width="15%">
                            $($_.'UsedSpace_MB')
                            </td>
                            <td width="15%">
                            $($_.'FreeSpace_MB')
                            </td>
                            <td width="15%">
                            $($_.'GrowthSpace_MB')
                            </td>
                        </tr>
                        </table>
                                </td>
                </tr>
                </tbody>
                </table>
"@

                
  #@ -f $_.'DBName', $_.'FileSize_MB', $_.'UsedSpace_MB', $_.'FreeSpace_MB', $_.'GrowthSpace_MB')
}
#endregion

#######################################################################
#region HTML - Discoverd Systems with client & Active Clients
#######################################################################

# Set html
$html = $html + @"
<table width="930" border="1">
  <tbody>
    <tr>
      <td><table width="400">
        <tr><h4> Discovered Systems with Client Installed</h4></tr>
        <tr>
          <td style="background-color:$(Set-PercentageColour -Value $Data.ClientCountPercentage);color:#ffffff;" width="$($Data.ClientCountPercentage)%"> $($Data.ClientCountPercentage)% </td>
          <td style="background-color:#eeeeee;color:#333333;" width="$($Data.NoClientCountPercentage)%"></td>
        </tr>
      </table>
        <table width="400">
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
        <table width="400">
          <tr>
            <td style="background-color:$(Set-PercentageColour -Value $Data.ActiveCountPercentage);color:#ffffff;" width="$($Data.ActiveCountPercentage)%"> $($Data.ActiveCountPercentage)% </td>
            <td style="background-color:#eeeeee;color:#333333;" width="$($Data.InactiveCountPercentage)%"></td>
          </tr>
        </table>
        <table width="400">
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
<table width="930" border="1">
  <tbody>
    <tr>
      <td><h4>Active Clients Health Evaluation</h4>
        <table width="400">
            <tr>
                <td style="background-color:$(Set-PercentageColour -Value $Data.ActivePassCountPercentage);color:#ffffff;" width="$($Data.ActivePassCountPercentage)%">
                $($Data.ActivePassCountPercentage)%
                </td>
                <td style="background-color:#eeeeee;color:#333333;" width="$($Data.ActiveNotPassCountPercentage)%">
                </td>
            </tr>
        </table>
        <table width="400">
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
        <table width="400">
          <tr>
            <td style="background-color:$(Set-PercentageColour -Value $Data.ActiveCountPercentage);color:#ffffff;" width="$($Data.ActiveCountPercentage)%"> $($Data.ActiveCountPercentage)% </td>
            <td style="background-color:#eeeeee;color:#333333;" width="$($Data.InactiveCountPercentage)%"></td>
          </tr>
        </table>
        <table width="400">
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
#region HTML Hardware Inventory & Software Inventory
#######################################################################

# Set html
$html = $html + @"
<table width="930" border="1">
  <tbody>
    <tr>
        <td><h4>Active Clients Hardware Inventory</h4>
      <table width="400">
        <tr>
          <td style="background-color:$(Set-PercentageColour -Value $Data.ActiveHWCountPercentage -UseInventoryThresholds);color:#ffffff;" width="$($Data.ActiveHWCountPercentage)%">
          $($Data.ActiveHWCountPercentage)%
          </td>
          <td style="background-color:#eeeeee;color:#333333;" width="$($Data.InActiveHWCountPercentage)%">
          </td>
        </tr>
      </table>
      <table width="400">
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
      <table width="400">
        <tr>
          <td style="background-color:$(Set-PercentageColour -Value $Data.ActiveSWCountPercentage -UseInventoryThresholds);color:#ffffff;" width="$($Data.ActiveSWCountPercentage)%">
          $($Data.ActiveSWCountPercentage)%
          </td>
          <td style="background-color:#eeeeee;color:#333333;" width="$($Data.InActiveSWCountPercentage)%">
          </td>
        </tr>
      </table>
      <table width="400">
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
#region HTML Active PolicyRequest Count
#######################################################################

# Set html
$html = $html + @"
    <table width="930" border="1" bordercolor="black">
    <tbody>
        <tr>
            <td width="400">
                <table cellspacing="0">
                    <tr><h4>Active Clients Policy Request</h4></tr>
                </table>
                <table width="400">
                    <tr>
                        <td style="background-color:$(Set-PercentageColour -Value $Data.ActivePRCountPercentage);color:#ffffff;" width="$($Data.ActivePRCountPercentage)%">
                        $($Data.ActivePRCountPercentage)%
                        </td>
                        <td style="background-color:#eeeeee;color:#333333;" width="$($Data.InActivePRCountPercentage)%">
                        </td>
                    </tr>
                </table>
                <table width="400">
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
