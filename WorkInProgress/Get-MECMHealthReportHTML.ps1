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
$StatusMessageTime = (Get-Date).AddDays(-2)
# Number of Status messages to report
$SMCount = 5
# Tally interval - see https://docs.microsoft.com/en-us/sccm/develop/core/servers/manage/about-configuration-manager-tally-intervals
$TallyInterval = '0001128000100008'
# Location of the resource dlls in the SCCM admin console path
$script:SMSMSGSLocation = “$env:SMS_ADMIN_UI_PATH\00000409”

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
    $Hex = '#90D7A5' # Green
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

function Get-StatusMessage {
    param (
        $MessageID,
        [ValidateSet("srvmsgs.dll","provmsgs.dll","climsgs.dll")]$DLL,
        [ValidateSet("Informational","Warning","Error")]$Severity,
        $InsString1,
        $InsString2,
        $InsString3,
        $InsString4,
        $InsString5,
        $InsString6,
        $InsString7,
        $InsString8,
        $InsString9,
        $InsString10
    )

    # Set the resources dll
    Switch ($DLL)
    {
        "srvmsgs.dll" { $stringPathToDLL = "$SMSMSGSLocation\srvmsgs.dll" }
        "provmsgs.dll" { $stringPathToDLL = "$SMSMSGSLocation\provmsgs.dll" }
        "climsgs.dll" { $stringPathToDLL = "$SMSMSGSLocation\climsgs.dll" }
    }

    # Load Status Message Lookup DLL into memory and get pointer to memory 
    $ptrFoo = $Win32LoadLibrary::LoadLibrary($stringPathToDLL.ToString()) 
    $ptrModule = $Win32GetModuleHandle::GetModuleHandle($stringPathToDLL.ToString()) 
    
    # Set severity code
    Switch ($Severity)
    {
        "Informational" { $code = 1073741824 }
        "Warning" { $code = 2147483648 }
        "Error" { $code = 3221225472 }
    }

    # Format the message
    $result = $Win32FormatMessage::FormatMessage($flags, $ptrModule, $Code -bor $MessageID, 0, $stringOutput, $sizeOfBuffer, $stringArrayInput)
    if ($result -gt 0)
        {
            # Add insert strings to message
            $objMessage = New-Object System.Object 
            $objMessage | Add-Member -type NoteProperty -name MessageString -value $stringOutput.ToString().Replace("%11","").Replace("%12","").Replace("%3%4%5%6%7%8%9%10","").Replace("%1",$InsString1).Replace("%2",$InsString2).Replace("%3",$InsString3).Replace("%4",$InsString4).Replace("%5",$InsString5).Replace("%6",$InsString6).Replace("%7",$InsString7).Replace("%8",$InsString8).Replace("%9",$InsString9).Replace("%10",$InsString10)
        }

    Return $objMessage
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
# Powershell - Clienthealth summary
#######################################
$clientHealthSummary = Get-CMClientHealthSummary -CollectionName 'All systems'  

$DeviceWithoutClient = get-CMDevice | Where-Object { $_.IsClient -eq $false } | Sort-Object name | Select-Object Name

$clientHealthSummaryWithoutClient = $clientHealthSummary.ClientsTotal - $DeviceWithoutClient.Count

$clientHealthSummarypercentage = [Math]::Round($clientHealthSummary.ClientsHealthy / $clientHealthSummaryWithoutClient * 100)
$clientHealthSummaryNOTPercentage = 100 - $clientHealthSummarypercentage

 
#endregion

#######################################################################
#region Query Region
#######################################################################
# Create has table to store data
$Data = @{}

###########################################
# QUERY - Overall status Site
###########################################
$query ="Select
SiteStatus.SiteCode, SiteInfo.SiteName, SiteStatus.Updated 'Time Stamp',
Case SiteStatus.Status
When 0 Then 'OK'
When 1 Then 'Warning'
When 2 Then 'Critical'
Else ' '
End AS 'Site Status',
Case SiteInfo.Status
When 1 Then 'Active'
When 2 Then 'Pending'
When 3 Then 'Failed'
When 4 Then 'Deleted'
When 5 Then 'Upgrade'
Else ' '
END AS 'Site State'
From V_SummarizerSiteStatus SiteStatus Join v_Site SiteInfo on SiteStatus.SiteCode = SiteInfo.SiteCode
Order By SiteCode" 

$data.Sitestatus = Get-SQLData -Query $query

###########################################
# QUERY - Component Status
###########################################

# SQL query for component status
$Query = "
Select 
	ComponentName,
	ComponentType,
	Case
		when Status = 0 then 'OK'
		when Status = 1 then 'Warning'
		when Status = 2 then 'Critical'	
	End as 'Status',
	Case
		when State = 0 then 'Stopped'
		when State = 1 then 'Started'
		when State = 2 then 'Paused'
		when State = 3 then 'Installing'
		when State = 4 then 'Re-installing'
		when State = 5 then 'De-installing'
	End as 'State',
	Case
		When AvailabilityState = 0 then 'Online'
		When AvailabilityState = 3 then 'Offline'
		When AvailabilityState = 4 then 'Unknown'
	End as 'AvailabilityState',
	Infos,
	Warnings,
	Errors
from vSMS_ComponentSummarizer
where TallyInterval = N'$TallyInterval'
and MachineName = '$SiteServer'
and SiteCode = '$SiteCode '
and Status in (1,2)
Order by Status,ComponentName
"
$data.ComponentStatus = Get-SQLData -Query $Query


###########################################
# QUERY - Disk and SQL
###########################################

$query = "SELECT distinct
Case v_SiteSystemSummarizer.Status
When 0 Then 'OK'
When 1 Then 'Warning'
When 2 Then 'Critical'
Else ' '
End As 'Status',
SiteCode 'Site Code',
SUBSTRING(SiteSystem, CHARINDEX('\\', SiteSystem) + 2, CHARINDEX(']', SiteSystem) - CHARINDEX('\\', SiteSystem) - 3 ) AS 'Site System',REPLACE(Role, 'SMS', 'ConfigMgr') 'Role',
SUBSTRING(SiteObject, CHARINDEX('Display=', SiteObject) + 8, CHARINDEX(']', SiteObject) - CHARINDEX('Display=',SiteObject) - 9) AS 'Storage Object',
Case ObjectType
When 0 Then 'Directory'
When 1 Then 'SQL Database'
When 2 Then 'SQL Transaction Log'
Else ' '
END AS 'Object Type',
CAST(BytesTotal/1024 AS VARCHAR(49)) + 'MB' 'Total',
CAST(BytesFree/1024 AS VARCHAR(49)) + 'MB' 'Free',
CASE PercentFree
When -1 Then 'Unknown'
When -2 Then 'Automatically grow'
ELSE CAST(PercentFree AS VARCHAR(49)) + '%'
END AS '%Free'
FROM v_SiteSystemSummarizer
Order By 'Storage Object'"

$Data.DiskSiteSQL = Get-SQLData -Query $Query

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
# Html CSS style
$Style = @"
<style>
table { 
    border-collapse: collapse;
    width: 930px;
}
td, th { 
    border: 1px solid #ddd;
    padding: 8px;
}
th {
    padding-top: 12px;
    padding-bottom: 12px;
    text-align: left;
    background-color: #4286f4;
    color: white;
}
h4 {
    color: Yellow;
    width:930px;
}
</style>
"@




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
      padding-left: 5px;
      padding-right: 5px;
      vertical-align:top;
      font-size:11.5px

      
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
#region HTML Overall Site Status
#######################################################################

# Convert results to HTML
$htmlData = $data.Sitestatus | 
    ConvertTo-Html -Property "Sitecode","SiteName","TimeStamp","Site Status","Site State" -Head $Style -Body "<h4>Overall Site Status</h4>" -CssUri "http://www.w3schools.com/lib/w3.css" | 
    Out-String
$HTML = $html + $htmlData 

#endregion

#######################################################################
#region HTML ADR Status
#######################################################################

# Convert results to HTML
$htmlData = $ADRstatus | 
    ConvertTo-Html -Property "Name","LastErrorCode","LastErrorTime","LastRunTime" -Head $Style -Body "<h4>ADR Status</h4>" -CssUri "http://www.w3schools.com/lib/w3.css" | 
    Out-String
$HTML = $html + $htmlData 

#endregion

#######################################################################
#region HTML ComponentStatus Siteserver
#######################################################################



# Convert results to HTML
$htmlData = $data.ComponentStatus | 
    ConvertTo-Html -Property "ComponentName","ComponentType","Status","State","AvailabilityState","Infos","Warnings","Errors" -Head $Style -Body "<h4>Components in a Warning or Error State</h4>" -CssUri "http://www.w3schools.com/lib/w3.css" | 
    Out-String
$HTML = $html + $htmlData + "<h4>Last $SMCount Error or Warning Status Messages for...</h4>" 


If ($data.ComponentStatus)
{

    # Start PInvoke Code 
$sigFormatMessage = @' 
[DllImport("kernel32.dll")] 
public static extern uint FormatMessage(uint flags, IntPtr source, uint messageId, uint langId, StringBuilder buffer, uint size, string[] arguments); 
'@ 
 
$sigGetModuleHandle = @' 
[DllImport("kernel32.dll")] 
public static extern IntPtr GetModuleHandle(string lpModuleName); 
'@ 
 
$sigLoadLibrary = @' 
[DllImport("kernel32.dll")] 
public static extern IntPtr LoadLibrary(string lpFileName); 
'@ 
 
    $Win32FormatMessage = Add-Type -MemberDefinition $sigFormatMessage -name "Win32FormatMessage" -namespace Win32Functions -PassThru -Using System.Text 
    $Win32GetModuleHandle = Add-Type -MemberDefinition $sigGetModuleHandle -name "Win32GetModuleHandle" -namespace Win32Functions -PassThru -Using System.Text 
    $Win32LoadLibrary = Add-Type -MemberDefinition $sigLoadLibrary -name "Win32LoadLibrary" -namespace Win32Functions -PassThru -Using System.Text 
    #End PInvoke Code 
 
    $sizeOfBuffer = [int]16384 
    $stringArrayInput = {"%1","%2","%3","%4","%5", "%6", "%7", "%8", "%9"} 
    $flags = 0x00000800 -bor 0x00000200  
    $stringOutput = New-Object System.Text.StringBuilder $sizeOfBuffer 

    # Process each resulting component
    Foreach ($Result in $data.ComponentStatus)
    {
        # Query SQL for status messages 
        $Component = $Result.ComponentName
        $SMQuery = "
        select 
	        top $SMCount
	        smsgs.RecordID, 
	        CASE smsgs.Severity 
		        WHEN -1073741824 THEN 'Error' 
		        WHEN 1073741824 THEN 'Informational' 
		        WHEN -2147483648 THEN 'Warning' 
		        ELSE 'Unknown' 
	        END As 'SeverityName', 
	        case smsgs.MessageType
		        WHEN 256 THEN 'Milestone'
		        WHEN 512 THEN 'Detail'
		        WHEN 768 THEN 'Audit'
		        WHEN 1024 THEN 'NT Event'
		        ELSE 'Unknown'
	        END AS 'Type',
	        smsgs.MessageID, 
	        smsgs.Severity, 
	        smsgs.MessageType, 
	        smsgs.ModuleName,
	        modNames.MsgDLLName, 
	        smsgs.Component, 
	        smsgs.MachineName, 
	        smsgs.Time, 
	        smsgs.SiteCode, 
	        smwis.InsString1, 
	        smwis.InsString2, 
	        smwis.InsString3, 
	        smwis.InsString4, 
	        smwis.InsString5, 
	        smwis.InsString6, 
	        smwis.InsString7, 
	        smwis.InsString8, 
	        smwis.InsString9, 
	        smwis.InsString10  
        from v_StatusMessage smsgs   
        join v_StatMsgWithInsStrings smwis on smsgs.RecordID = smwis.RecordID
        join v_StatMsgModuleNames modNames on smsgs.ModuleName = modNames.ModuleName
        where smsgs.MachineName = '$SiteServer' 
        and smsgs.Component = '$Component'
        and smsgs.Severity in ('-1073741824','-2147483648')
        Order by smsgs.Time DESC
        "
        $StatusMsgs = Get-SQLData -Query $SMQuery

        # Put desired fields into an object for each result
        $StatusMessages = @()
        foreach ($Row in $StatusMsgs)
        {
            $Params = @{
                MessageID = $Row.MessageID
                DLL = $Row.MsgDLLName
                Severity = $Row.SeverityName
                InsString1 = $Row.InsString1
                InsString2 = $Row.InsString2
                InsString3 = $Row.InsString3
                InsString4 = $Row.InsString4
                InsString5 = $Row.InsString5
                InsString6 = $Row.InsString6
                InsString7 = $Row.InsString7
                InsString8 = $Row.InsString8
                InsString9 = $Row.InsString9
                InsString10 = $Row.InsString10
                }
            $Message = Get-StatusMessage @params

            $StatusMessage = New-Object psobject
            Add-Member -InputObject $StatusMessage -Name Severity -MemberType NoteProperty -Value $Row.SeverityName
            Add-Member -InputObject $StatusMessage -Name Type -MemberType NoteProperty -Value $Row.Type
            Add-Member -InputObject $StatusMessage -Name SiteCode -MemberType NoteProperty -Value $Row.SiteCode
            Add-Member -InputObject $StatusMessage -Name "Date / Time" -MemberType NoteProperty -Value $Row.Time
            Add-Member -InputObject $StatusMessage -Name System -MemberType NoteProperty -Value $Row.MachineName
            Add-Member -InputObject $StatusMessage -Name Component -MemberType NoteProperty -Value $Row.Component
            Add-Member -InputObject $StatusMessage -Name Module -MemberType NoteProperty -Value $Row.ModuleName
            Add-Member -InputObject $StatusMessage -Name MessageID -MemberType NoteProperty -Value $Row.MessageID
            Add-Member -InputObject $StatusMessage -Name Description -MemberType NoteProperty -Value $Message.MessageString
            $StatusMessages += $StatusMessage
        }

        # Add to the HTML code
        $HTML = $HTML + (
            $StatusMessages | 
                ConvertTo-Html -Property "Severity","Date / Time","MessageID","Description" -Head $Style -Body "<h4>$Component</h4>" -CssUri "http://www.w3schools.com/lib/w3.css" | 
                Out-String
            )

    }
}

#endregion

#######################################################################
#region HTML Disk Site SQL Siteserver
#######################################################################

# Convert results to HTML
$htmlData = $data.DiskSiteSQL | 
    ConvertTo-Html -Property "Status","Site System","Role","ObjectType","Total","Free","%Free" -Head $Style -Body "<h4>Disk & SQL Status</h4>" -CssUri "http://www.w3schools.com/lib/w3.css" | 
    Out-String
$HTML = $html + $htmlData 

#endregion

#######################################################################
#region HTML Client Versions
#######################################################################

# Convert results to HTML
$htmlData = $data.ClientVersions | 
    ConvertTo-Html -Property "Version","Count","Percent" -Head $Style -Body "<h4>Client Version</h4>" -CssUri "http://www.w3schools.com/lib/w3.css" | 
    Out-String
$HTML = $html + $htmlData 
    
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

$lastdate = (Get-Date).AddDays(-7)

# Set html
$html = $html + @"
    <table width="930" border="1">
    <tbody>
    <tr>
        <td>
            <h4>Maintenance Task Status Last 24 hours since $lastdate</h4>
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
  
if ($_.'lastCompletionTime' -gt $lastdate)
{
    
    $html = $html + @"
    <table width="930" border="1">
    <tbody>
    <tr>
        <td>
            <table width="100%" style="background-color: #90D7A5">
                <tr>

                    <td width="50%" style="background-color: #90D7A5">
                    $($_.'Taskname')
                    </td>

                         <td width="25%" style="background-color: #90D7A5">
                    $($_.'LastStartTime')
                    </td>
                            <td width="25%" style="background-color: #90D7A5">
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
else
{
   
   $html = $html + @"
    <table width="930" border="1">
    <tbody>
    <tr bgcolor="green">
        <td>
            <table width="100%" style="background-color: #E8B342">
                <tr>

                    <td width="50%" style="background-color: #E8B342">
                    $($_.'Taskname')
                    </td>

                         <td width="25%" style="background-color: #E8B342">
                    $($_.'LastStartTime')
                    </td>
                            <td width="25%" style="background-color: #E8B342">
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
  
  
  
}
#endregion

#######################################################################
#region HTML - Discoverd Systems with client & Active Clients Policty Request
#######################################################################

# Set html
$html = $html + @"
<table width="930" border="1">
  <tbody>
    <tr>
      <td width="50%"><table width="400">
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
            <td width="445">
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
#region HTML Summary Healthy Device
#######################################################################

# Set html
$html = $html + @"
<table width="930" border="1">
  <tbody>
    <tr>
        <td><h4>Summary Healthy Device</h4>
      <table width="400">
        <tr>
          <td style="background-color:$(Set-PercentageColour -Value $clientHealthSummarypercentage -UseInventoryThresholds);color:#ffffff;" width="$($clientHealthSummarypercentage)%">
          $($clientHealthSummarypercentage)%
          </td>
          <td style="background-color:#eeeeee;color:#333333;" width="$($clientHealthSummaryNOTPercentage)%">
          </td>
        </tr>
      </table>
      <table width="100%">
        <tr>
            <td width="80%">
            Device Total
            </td>
            <td width="20%">
            $($clientHealthSummary.ClientsTotal)
            </td>
        </tr>
        <tr>
            <td width="80%">
            Device Healthy
            </td>
            <td width="20%">
            $($clientHealthSummary.ClientsHealthy)
            </td>
        </tr>
        <tr>
            <td width="80%">
            Device Unhealthy
            </td>
            <td width="20%">
            $($clientHealthSummary.ClientsUnhealthy)
            </td>
        </tr>
        <tr>
            <td width="80%">
            Device Health unknown
            </td>
            <td width="20%">
            $($clientHealthSummary.ClientsHealthUnknown)
            </td>
        </tr>
        <tr>
            <td width="80%">
            Device without Client
            </td>
            <td width="20%">
            $($DeviceWithoutClient.Count)
            </td>
        </tr>
      </table>
      </td>
      <td><h4></h4>
      <table width="400">
        <tr>

        </tr>
      </table>
      <table width="400">
        <tr>

        </tr>
        <tr>

        </tr>
      </table>
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
