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
$sitecode = 'ps1:'
$siteserver = 'TH-mgt02.korsberga.local'
$StatusMessageTime = (Get-Date).AddDays(-2)
# Number of Status messages to report
$SMCount = 5
# Tally interval - see https://docs.microsoft.com/en-us/sccm/develop/core/servers/manage/about-configuration-manager-tally-intervals
$TallyInterval = '0001128000100008'
# Location of the resource dlls in the SCCM admin console path
$script:SMSMSGSLocation = “$env:SMS_ADMIN_UI_PATH\00000409”


# The no-reply emailaddress
$Emailfrom = 'no-reply@trivselhus.se'
#
# The email (group) who will receive the report
$email_Error = 'christian.damberg@trivselhus.se'
$emailto = 'christian.damberg@trivselhus.se'
#
# The email when the script cant find any updates
$email_noErrors = 'christian.damberg@trivselhus.se'
#
# SMTP-server
$smtp = 'webmail.trivselhus.se'

#endregion

#######################################################################
#region Functions
#######################################################################

# Check for ConfigMgr 
if (-not(Get-Module -name ConfigurationManager)) {
  Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')
}

# MAilkitMessage
if (-not(Get-Module -name send-mailkitmessage)) {
  Install-Module send-mailkitmessage
  Import-Module send-mailkitmessage
}

# EnhancedHTML2
if (-not(Get-Module -name EnhancedHTML2)) {
  Install-Module -Name EnhancedHTML2
  Import-Module EnhancedHTML2
}
#########################################################
# To run the script you must be on ps-drive for MEMCM
#########################################################
Push-Location
Set-Location $SiteCode

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
$datainbox = Get-WmiObject -Class Win32_PerfFormattedData_SMSINBOXMONITOR_SMSInbox -ComputerName $siteserver| Where-Object filecurrentcount -gt '0' | Select-Object -Property PSComputerName, Name, FileCurrentCount

$dataevents = get-eventlog system -After (Get-Date).AddDays(-7) -EntryType Error  

#endregion

#######################################################################
#Query Region
#######################################################################
# Create has table to store data
$Data = @{}
#######################################################################
# SYSTEM
#######################################################################

###########################################
#region QUERY - Overall status Site
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
#endregion
###########################################
#region QUERY - ADR Status
###########################################
$query ="
SELECT
Name, 
AutodeploymentEnabled, 
lastruntime, 
LastErrorcode,
CASE when lasterrorcode > '0' THEN 'Error' ELSE 'TASK Successful' END AS 'LastRun',

CASE when AutodeploymentEnabled = '1' Then 'Enabled' Else 'Disabled' END AS 'Status'
FROM vSMS_AutoDeployments
order by Name
" 
$data.ADRStatus = Get-SQLData -Query $query
#endregion
###########################################
#region QUERY - Component Status
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
#endregion
###########################################
#region QUERY - Disk and SQL
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
#endregion
############################################
#region QUERY - All Client Version
############################################ 

$Query = "
Declare @CollectionID Varchar(8)
Set @CollectionID = 'SMS00001'
Select sys.Client_version0 as 'Client Version',
CASE
WHEN client_version0 = '5.00.8412.1000'Then'MECM 1606'
WHEN client_version0 = '5.00.8853.1006'Then'MECM 1906'
WHEN client_version0 = '5.00.9012.1020'Then'MECM 2006'
WHEN client_version0 = '5.00.9068.1008'Then'MECM 2111'
WHEN client_version0 = '5.00.9068.1012'Then'MECM 2111 hotfix KB12959506'

ELSE
client_version0
END as 'ConfigMgr Release',
Count(DISTINCT sys.ResourceID) as 'Client Count',
(STR((COUNT(sys.ResourceID)*100.0/(
Select COUNT(SYS.ResourceID)
From v_FullCollectionMembership FCM INNER JOIN V_R_System sys on FCM.ResourceID = SYS.ResourceID
Where FCM.CollectionID = @CollectionID
and
Sys.Client0= '1')),5,2)) + ' %' AS 'Percent %'
From v_FullCollectionMembership FCM INNER JOIN V_R_System sys on FCM.ResourceID = SYS.ResourceID
Where SYS.Client0 = '1' and FCM.CollectionID = @CollectionID
Group By sys.Client_version0
Order by sys.Client_version0 DESC
"
$Data.Clientversion = Get-SQLData -Query $Query
#endregion
############################################
#region QUERY - Client Health Thresholds
############################################

$Query = '
  SELECT *
  FROM v_CH_Settings
  where SettingsID = 1
'

$Data.CHSettings = Get-SQLData -Query $Query 
#endregion
############################################
#region QUERY - Client Installation failure
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
#endregion
############################################
#region QUERY - All DP Status
############################################

$Query = "
select UPPER
(SUBSTRING(PSD.ServerNALPath,13,CHARINDEX('.', PSd.ServerNALPath) -13)) AS [DP Name],
count(*) [Targeted] ,
count(CASE when PSD.State='0' then '*' END) AS 'Installed',
count(CASE when PSD.State not in ('0') then '*' END) AS 'Not Installed',
round((CAST(SUM (CASE WHEN PSD.State='0' THEN 1 ELSE 0 END) as float)/COUNT(psd.PackageID ) )*100,2) as 'Success%',
psd.SiteCode [Reporting Site]
From v_PackageStatusDistPointsSumm psd,SMSPackages P
where p.PackageType!=4
and (p.PkgID=psd.PackageID)
group by PSd.ServerNALPath,psd.SiteCode
"

$Data.DPstatus = Get-SQLData -Query $Query
#endregion
############################################
#region QUERY - All Maintenance Task Status
############################################

# All Maintenance Task
$Query = "
 SELECT
TaskName,
LastStartTime,
LastCompletionTime,
CASE WHEN CompletionStatus = '1' THEN 'Task failed' ELSE 'Task successful' END AS 'Status'
FROM
dbo.SQLTaskStatus
WHERE
(NOT (LastStartTime LIKE CONVERT(DATETIME, '1980-01-01 00:00:00', 102)))
"

$Data.MWStatus = Get-SQLData -Query $Query

# Maintenance Task for Backup

$query = "
 SELECT
TaskName,
LastStartTime,
LastCompletionTime,
CASE WHEN CompletionStatus = '1' THEN 'Task failed' ELSE 'Task successful' END AS 'Status'
FROM
dbo.SQLTaskStatus
WHERE
(NOT (LastStartTime LIKE CONVERT(DATETIME, '1980-01-01 00:00:00', 102)))

AND TaskName Like 'Backup%'
"
$data.BackupStatus = Get-SQLData -Query $query

#endregion
############################################
#region QUERY - All Policy Request
############################################ 

$Query = "
  Declare @CollectionID as Varchar(8)
Declare @TotalActive as Numeric(8)
Declare @ActivePolicyRequest as Numeric(8)
Declare @InActivePolicyRequest as Numeric(8)
Set @CollectionID = 'SMS00001' --Specify the collection ID
select @TotalActive = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (Ch.ClientActiveStatus = 1))
)
select @ActivePolicyRequest = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (IsActivePolicyRequest = 1 and ClientActiveStatus = 1))
)
select @InActivePolicyRequest = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (IsActivePolicyRequest = 0 and ClientActiveStatus = 1))
)
select
@TotalActive as 'TotalActive',
@ActivePolicyRequest as 'ActivePolicyRequest',
@InActivePolicyRequest as 'InActivePolicyRequest',
case when (@TotalActive = 0) or (@TotalActive is null) Then '100' Else (round(@ActivePolicyRequest/ convert
(float,@TotalActive)*100,2)) End as 'ActivePolicyRequest%'

"
$Data.ActiveWorkstationPolicyRequestCount = Get-SQLData -Query $Query 
#endregion
###########################################
#region QUERY - All Active Client Heartbeat (DDR) Status
###########################################

$Query = "
  Declare @CollectionID as Varchar(8)
Declare @TotalActive as Numeric(8)
Declare @ActiveHeartBeatDDR as Numeric(8)
Declare @InActiveHeartBeatDDR as Numeric(8)
Set @CollectionID = 'SMS00001' --Specify the collection ID
select @TotalActive = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (Ch.ClientActiveStatus = 1))
)
select @ActiveHeartBeatDDR = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (IsActiveDDR = 1 and ClientActiveStatus = 1))
)
select @InActiveHeartBeatDDR = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (IsActiveDDR = 0 and ClientActiveStatus = 1))
)
select
@TotalActive as 'TotalActive',
@ActiveHeartBeatDDR as 'ActiveHeartBeatDDR',
@InActiveHeartBeatDDR as 'InActiveHeartBeatDDR',
case when (@TotalActive = 0) or (@TotalActive is null) Then '100' Else (round(@ActiveHeartBeatDDR/ convert
(float,@TotalActive)*100,2)) End as 'ActiveHeartBeatDDR%'

"
$Data.ActiveDDRWorkstationCount = Get-SQLData -Query $Query

#endregion
###########################################
#region QUERY - All Active Client Hardware Inventory Status
###########################################
$Query = "
 Declare @CollectionID as Varchar(8)
Declare @TotalActive as Numeric(8)
Declare @ActiveHWInv as Numeric(8)
Declare @InActiveHWInv as Numeric(8)
Set @CollectionID = 'SMS00001' --Specify the collection ID
select @TotalActive = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (Ch.ClientActiveStatus = 1))
)
select @ActiveHWInv = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (IsActiveHW = 1 and ClientActiveStatus = 1))
)
select @InActiveHWInv = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (IsActiveHW = 0 and ClientActiveStatus = 1))
)
select
@TotalActive as 'TotalActive',
@ActiveHWInv as 'ActiveHWInv',
@InActiveHWInv as 'InActiveHWInv',
case when (@TotalActive = 0) or (@TotalActive is null) Then '100' Else (round(@ActiveHWInv/ convert (float,@TotalActive)*100,2))
End as 'ActiveHWInv%'
"
$Data.ActiveHardWareInventoryWorkstationCount = Get-SQLData -Query $Query

#endregion
###########################################
#region QUERY - All Active Client Hardware Inventory Status
###########################################

$Query = "
 Declare @CollectionID as Varchar(8)
Declare @TotalActive as Numeric(8)
Declare @ActiveHWInv as Numeric(8)
Declare @InActiveHWInv as Numeric(8)
Set @CollectionID = 'SMS00001' --Specify the collection ID
select @TotalActive = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (Ch.ClientActiveStatus = 1))
)
select @ActiveHWInv = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (IsActiveHW = 1 and ClientActiveStatus = 1))
)
select @InActiveHWInv = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (IsActiveHW = 0 and ClientActiveStatus = 1))
)
select
@TotalActive as 'TotalActive',
@ActiveHWInv as 'ActiveHWInv',
@InActiveHWInv as 'InActiveHWInv',
case when (@TotalActive = 0) or (@TotalActive is null) Then '100' Else (round(@ActiveHWInv/ convert (float,@TotalActive)*100,2))
End as 'ActiveHWInv%'
"
$Data.ActiveHardWareInventoryWorkstationCount = Get-SQLData -Query $Query

#endregion
###########################################
#region QUERY - All Active Client Software Inventory Status
###########################################

$Query = "
Declare @CollectionID as Varchar(8)
Declare @TotalActive as Numeric(8)
Declare @ActiveSWInv as Numeric(8)
Declare @InActiveSWInv as Numeric(8)
Set @CollectionID = 'SMS00001' --Specify the collection ID
select @TotalActive = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (Ch.ClientActiveStatus = 1))
)
select @ActiveSWInv = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (IsActiveSW = 1 and ClientActiveStatus = 1))
)
select @InActiveSWInv = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (IsActiveSW = 0 and ClientActiveStatus = 1))
)
select
@TotalActive as 'TotalActive',
@ActiveSWInv as 'ActiveSWInv',
@InActiveSWInv as 'InActiveSWInv',
case when (@TotalActive = 0) or (@TotalActive is null) Then '100' Else (round(@ActiveSWInv/ convert (float,@TotalActive)*100,2))
End as 'ActiveSWInv%'
"
$Data.ActiveSoftwareInventoryWorkstationCount = Get-SQLData -Query $Query

#endregion
############################################
#region QUERY - All With Client
############################################ 

$Query = "
Declare @CollectionID as Varchar(8)
Declare @TotalSystem as Numeric(8)
Declare @WithClient as Numeric(8)
Declare @NoClient as Numeric(8)
Set @CollectionID = 'SMS00001' --Specify the collection ID

select @TotalSystem = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_R_System Ch on Vrs.ResourceID = ch.ResourceID
where (Ch.Client0 = 1 or ch.Client0 = 0 or ch.client0 is null))
)

select @WithClient = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_R_System Ch on Vrs.ResourceID = ch.ResourceID
where (Ch.Client0 = 1))
)

select @NoClient = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_R_System Ch on Vrs.ResourceID = ch.ResourceID
where (ch.Client0 = 0 or ch.client0 is null))
)

select
@TotalSystem as 'TotalSystem',
@WithClient as 'WithClient',
@NoClient as 'NoClient',
case when (@TotalSystem = 0) or (@TotalSystem is null) Then '100' Else (round(@WithClient/
convert (float,@TotalSystem)*100,2)) End as 'WithClient%'



"
$Data.WorkstationClientCount = Get-SQLData -Query $Query 
#endregion
###########################################
#region QUERY - All Active Clients Health Evaluation Status
###########################################

$Query = "
Declare @CollectionID as Varchar(8)
Declare @TotalActive as Numeric(8)
Declare @Active_Pass as Numeric(8)
Declare @Active_Fail as Numeric(8)
Declare @Active_Unknown as Numeric(8)
Set @CollectionID = 'SMS00001' --Specify the collection ID

select @TotalActive = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (Ch.ClientActiveStatus = 1))
)

select @Active_Pass = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (ClientActiveStatus = 1 and ClientState = 1))
)

select @Active_Fail = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (ClientActiveStatus = 1 and ClientState = 2))
)

select @Active_Unknown = (
select COUNT(*) as 'Count' from v_FullCollectionMembership where CollectionID = @CollectionID 
and v_FullCollectionMembership.ResourceID in (
Select Vrs.ResourceID from v_R_System Vrs
inner join v_CH_ClientSummary Ch on Vrs.ResourceID = ch.ResourceID
where (ClientActiveStatus = 1 and ClientState = 3))
)
select
@TotalActive as 'TotalActive',
@Active_Pass as 'Active_Pass',
@Active_Fail as 'Active_Fail',
@Active_Unknown as 'Active_Unknown',
case when (@TotalActive = 0) or (@TotalActive is null) Then '100' Else (round(@Active_pass/
convert (float,@TotalActive)*100,2)) End as 'Active_Pass%'
"
$Data.ActiveWorkstationHealthEvalutionCount = Get-SQLData -Query $Query
  
#endregion
###########################################
#region QUERY - All Active Clients Diskspace
###########################################

$Query = "
Declare @CollectionID as Varchar(8)
Declare @FreeSpace as Integer
Set @CollectionID = 'SMS00001' -- specify scope collection ID
Set @FreeSpace = '20000' -- specify MB Size
Select
distinct (Vrs.Name0) as 'Machine',
Vrs.AD_Site_Name0 as 'ADSiteName',
Vrs.User_Name0 as 'UserName',
USR.Mail0 as 'EMailID',
Os.Caption00 as 'OSName',
Csd.SystemType00 as 'OSType',
LD.DeviceID00 as 'Drive',
LD.FileSystem00 as 'FileSystem',
LD.Size00 / 1024 as 'TotalSpace (GB)',
LD.FreeSpace00 / 1024 as 'FreeSpace (GB)',
Ws.LastHWScan as 'LastHWScan',
DateDiff(D, Ws.LastHwScan, GetDate()) as 'LastHWScanAge'
FROM v_R_System Vrs
Join v_R_User USR on USR.User_Name0 = Vrs.User_Name0
Join v_FullCollectionMembership Fc on Fc.ResourceID = Vrs.ResourceID
Join Operating_System_DATA Os on Os.MachineID = Vrs.ResourceID
Join Computer_System_DATA Csd on Csd.MachineID = Vrs.ResourceID
Join Logical_Disk_Data Ld on Ld.MachineID = Vrs.ResourceID
Join v_GS_WORKSTATION_STATUS Ws on Ws.ResourceID = Vrs.ResourceId
where CollectionID = @CollectionID
and LD.Description00 = 'Local Fixed Disk'
and LD.FreeSpace00 < @FreeSpace
and ld.DeviceID00 like 'c:'
Order By Vrs.Name0 asc
"
$Data.ClientDiskSpace = Get-SQLData -Query $Query
  
#endregion
############################################
#region QUERY - Application Deployment
############################################ 
$Query = "
Declare @CurrentDeploymentsReportNeededDays as integer
Set @CurrentDeploymentsReportNeededDays = 30 --Specify the Days
Select
CONVERT(VARCHAR(11),GETDATE(),106) as 'Date',
Right(Ds.CollectionName,3) as 'Stage',
Vaa.ApplicationName as 'ApplicationName',
CASE when Vaa.DesiredConfigType = 1 Then 'Install' when vaa.DesiredConfigType = 2 Then 'Uninstall' Else
'Others' End as 'DepType',
Ds.CollectionName as 'CollectionName',
CASE when Ds.DeploymentIntent = 1 Then 'Required' when Ds.DeploymentIntent = 2 Then 'Available' End as
'Purpose',
Ds.DeploymentTime as 'AvailableTime',
Ds.EnforcementDeadline as 'RequiredTime',
Ds.NumberTotal as 'Target',
Ds.NumberSuccess as 'Success',
Ds.NumberInProgress as 'Progress',
Ds.NumberErrors as 'Errors',
Ds.NumberOther as 'ReqNotMet',
Ds.NumberUnknown as 'Unknown',
case when (Ds.NumberTotal = 0) or (Ds.NumberTotal is null) Then '100' Else (round( (Ds.NumberSuccess +
Ds.NumberOther) / convert (float,Ds.NumberTotal)*100,2)) End as 'Success%',
DateDiff(D,Ds.EnforcementDeadline, GetDate()) as 'ReqDays'
from v_DeploymentSummary Ds
left join v_ApplicationAssignment Vaa on Ds.AssignmentID = Vaa.AssignmentID
Where Ds.FeatureType = 1 and Ds.DeploymentIntent = 1
and DateDiff(D,Ds.EnforcementDeadline, GetDate()) between 0 and @CurrentDeploymentsReportNeededDays
and Ds.NumberTotal > 0
order by Ds.EnforcementDeadline desc

"
$Data.ApplicationDeployment = Get-SQLData -Query $Query 
#endregion
############################################
#region QUERY - Package Deployment
############################################ 
$Query = "
Declare @CurrentDeploymentsReportNeededDays as integer
Set @CurrentDeploymentsReportNeededDays = 30 --Specify the Days
Select
CONVERT(VARCHAR(11),GETDATE(),106) as 'Date',
Right(Ds.CollectionName,3) as 'Stage',
Left(Ds.SoftwareName, CharIndex('(',(Ds.SoftwareName))-1)as 'ApplicationName',
Ds.ProgramName 'DepType',
Ds.CollectionName as 'CollectionName',
CASE when Ds.DeploymentIntent = 1 Then 'Required' when Ds.DeploymentIntent = 2 Then 'Available' End as
'Purpose',
Ds.DeploymentTime as 'AvailableTime',
Ds.EnforcementDeadline as 'RequiredTime',
Ds.NumberTotal as 'Target',
Ds.NumberSuccess as 'Success',
Ds.NumberInProgress as 'Progress',
Ds.NumberErrors as 'Errors',
Ds.NumberOther as 'ReqNotMet',
Ds.NumberUnknown as 'Unknown',
case when (Ds.NumberTotal = 0) or (Ds.NumberTotal is null) Then '100' Else (round( (Ds.NumberSuccess +
Ds.NumberOther) / convert (float,Ds.NumberTotal)*100,2)) End as 'Success%',
DateDiff(D,Ds.DeploymentTime, GetDate()) as 'AvailDays'
from v_DeploymentSummary Ds
join v_Advertisement Vaa on Ds.OfferID = Vaa.AdvertisementID
Where Ds.FeatureType = 2 and Ds.DeploymentIntent = 1
and DateDiff(D,Ds.DeploymentTime, GetDate()) between 0 and @CurrentDeploymentsReportNeededDays
and Ds.NumberTotal > 0
order by Ds.DeploymentTime desc
"
$Data.PackageDeployment = Get-SQLData -Query $Query 

#endregion

#######################################################################
#region Create html header
#######################################################################
# Html CSS style
$style = @"
<style>
body {
    color:#333333;
    font-family:Calibri,Tahoma;
    font-size: 10pt;
}

h1 {
    text-align:center;
}

h2 {
    border-top:1px solid #666666;
}
h4 {
    border-top:1px solid #666666;
}

th {
    font-weight:bold;
    color:#eeeeee;
    background-color:#333333;
    cursor:pointer;
}

.odd  { background-color:#ffffff; }

.even { background-color:#dddddd; }

.ok { background-color:lightgreen; }

.warning { background-color:lightyellow; }

.error { background-color:lightred; }

.paginate_enabled_next, .paginate_enabled_previous {
    cursor:pointer; 
    border:1px solid #222222; 
    background-color:#dddddd; 
    padding:2px; 
    margin:4px;
    border-radius:2px;
}

.paginate_disabled_previous, .paginate_disabled_next {
    color:#666666; 
    cursor:pointer;
    background-color:#dddddd; 
    padding:2px; 
    margin:4px;
    border-radius:2px;
}

.dataTables_info { margin-bottom:4px; }

.sectionheader { cursor:pointer; }

.sectionheader:hover { color:red; }

.grid { width:100% }

.red {
    color:red;
    font-weight:bold;
} 
</style>
"@

#########################################################
# The top of the email
#########################################################
$pre = @"
<br>
<img src='cid:logo.png' height="50">
<br>
<p><b>New updates!</b><br> 
<p>Updates will be available from wednesday week $weeknumber kl.15.00</p>
<p><b>Schema</b><br>
<p>The updates will be installed as follows:</p>
<p><ol>Test - Week $weeknumber - Every night between 03.00 - 08:00 (If any updates are published)</ol></p>
<p><ol>Prod - Week $nextweeknumber - Majority will be installed saturday 11.00pm till Sunday 09.00am</ol></p>
<p><ol>AX - Managed manually by the administration.</ol></p>
<p><b>Patchar From Microsoft</b><br>
<p>The following updates are downloaded and published in updategroup <b><i>$UpdateGroupName</i></b> since $limit</p>
<p>$Numbersofupdates</p>
"@

#########################################################
# Footer of the email
#########################################################
$post = @"
<p>Raport created $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b></p>
<p>Script created by:<br><a href="mailto:Your Email">Your name</a><br>
<a href="https://your blog">your description of your blog</a>
"@

$html = $pre

#endregion

#######################################################################
#region HTML Overall Site Status
#######################################################################

if ($data.Sitestatus.'Site Status' -eq 'OK')
{
           $params = @{'As'='Table';
                    'PreContent'='<h4>&diams; Site Status</h4>';
                    'EvenRowCssClass'='ok';
                    'OddRowCssClass'='ok';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';} 
}

if ($data.Sitestatus.'Site Status' -eq 'Warning')
{
           $params = @{'As'='Table';
                    'PreContent'='<h4>&diams; Site Status</h4>';
                    'EvenRowCssClass'='warning';
                    'OddRowCssClass'='warning';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';} 
}

if ($data.Sitestatus.'Site Status' -eq 'Error')
{
           $params = @{'As'='Table';
                    'PreContent'='<h4>&diams; Site Status</h4>';
                    'EvenRowCssClass'='error';
                    'OddRowCssClass'='error';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';} 
}        


        $htmlData = $data.sitestatus |
                   ConvertTo-EnhancedHTMLFragment @params
$HTML = $html + $htmlData 

#endregion
#######################################################################
#region HTML Maintenance Task status
#######################################################################

# Convert results to HTML

if ($data.BackupStatus.Status -eq 'TAsk Successful')
{
           $params = @{'As'='Table';
                    'PreContent'='<h4>&diams; Backup Status</h4>';
                    'EvenRowCssClass'='ok';
                    'OddRowCssClass'='ok';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';} 
}

if ($data.BackupStatus.Status -eq 'Task Failed')
{
           $params = @{'As'='Table';
                    'PreContent'='<h4>&diams; Backup Status</h4>';
                    'EvenRowCssClass'='warning';
                    'OddRowCssClass'='warning';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';} 
}

        $htmlData = $data.Backupstatus| ConvertTo-EnhancedHTMLFragment @params -Properties TaskName,LastStartTime,LastCompletionTime,Status
        

$HTML = $html + $htmlData





#endregion
#######################################################################
#region HTML All ADR Status
#######################################################################


# Convert results to HTML


        $params = @{'As'='Table';
                    'PreContent'='<h4>&diams; Automatic Deployment Rules</h4>';
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $data.ADRStatus |
                   ConvertTo-EnhancedHTMLFragment @params -Properties name,Status,AutodeploymentEnabled,Lastruntime,Lasterrorcode,Lastrun
        $HTML = $html + $htmlData 

#endregion
#######################################################################
#region HTML All ComponentStatus Siteserver
#######################################################################

If ($data.ComponentStatus)
{

# Convert results to HTML

        $params = @{'As'='Table';
                    'PreContent'='<h4>&diams; Component Status</h4>';
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $data.Componentstatus |
                   ConvertTo-EnhancedHTMLFragment @params

        $HTML = $html + $htmlData


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
            #Add-Member -InputObject $StatusMessage -Name Type -MemberType NoteProperty -Value $Row.Type
            #Add-Member -InputObject $StatusMessage -Name SiteCode -MemberType NoteProperty -Value $Row.SiteCode
            Add-Member -InputObject $StatusMessage -Name "Date / Time" -MemberType NoteProperty -Value $Row.Time
            #Add-Member -InputObject $StatusMessage -Name System -MemberType NoteProperty -Value $Row.MachineName
            Add-Member -InputObject $StatusMessage -Name Component -MemberType NoteProperty -Value $Row.Component
            #Add-Member -InputObject $StatusMessage -Name Module -MemberType NoteProperty -Value $Row.ModuleName
            #Add-Member -InputObject $StatusMessage -Name MessageID -MemberType NoteProperty -Value $Row.MessageID
            Add-Member -InputObject $StatusMessage -Name Description -MemberType NoteProperty -Value $Message.MessageString
            $StatusMessages += $StatusMessage
        }

        
        
        # Add to the HTML code

                $params = @{'As'='Table';
                    'PreContent'="<h4>&diams; Status Messages Last:$SMCount </h4>";
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $StatusMessages |
                   ConvertTo-EnhancedHTMLFragment @params

$HTML = $html + $htmlData


    }
}

#endregion
#######################################################################
#region HTML All Disk Site SQL Siteserver
#######################################################################

# Convert results to HTML

                $params = @{'As'='Table';
                    'PreContent'="<h4>&diams; Disk & SQL status </h4>";
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $data.Disksitesql |
                   ConvertTo-EnhancedHTMLFragment @params -Properties status,'site system',"Object Type",Total,Free,%Free


$HTML = $html + $htmlData 

#endregion
#######################################################################
#region HTML All Client Versions
#######################################################################

# Convert results to HTML

                $params = @{'As'='Table';
                    'PreContent'="<h4>&diams; Client Versions </h4>";
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $data.Clientversion |
                   ConvertTo-EnhancedHTMLFragment @params -Properties "Client Version","ConfigMgr Release","Client Count","Percent %"

$HTML = $html + $htmlData 
    

#endregion
#######################################################################
#region HTML Windows Client Installation Failures
#######################################################################

if ($data.InstallFailures)
{

                $params = @{'As'='Table';
                    'PreContent'="<h4>&diams; Client Installation Status (Failures) </h4>";
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $data.InstallFailures |
                   ConvertTo-EnhancedHTMLFragment @params

$HTML = $html + $htmlData 
}

#endregion
#######################################################################
#region HTML DP Status
#######################################################################

# Convert results to HTML

                $params = @{'As'='Table';
                    'PreContent'="<h4>&diams; Distribution Point Status </h4>";
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $data.DPstatus |
                   ConvertTo-EnhancedHTMLFragment @params

$HTML = $html + $htmlData 

#endregion
#######################################################################
#region HTML Inbox Status
#######################################################################

# Convert results to HTML
                $params = @{'As'='Table';
                    'PreContent'="<h4>&diams; Inbox Status </h4>";
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $datainbox |
                   ConvertTo-EnhancedHTMLFragment @params -Properties PScomputername,Name,FileCurrentCount


$HTML = $html + $htmlData 

#endregion
#######################################################################
#region HTML Events Status
#######################################################################

# Convert results to HTML

                $params = @{'As'='Table';
                    'PreContent'="<h4>&diams; System Events last 7 days</h4>";
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $dataevents |
                   ConvertTo-EnhancedHTMLFragment @params -Properties EntryType,Source,Message


$HTML = $html + $htmlData 

#endregion
#######################################################################
#region HTML All Client Status
#######################################################################

# Convert results to HTML

                $params = @{'As'='Table';
                    'PreContent'="<h4>&diams; All Workstation Client Status</h4>";
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $data.WorkstationClientCount |
                   ConvertTo-EnhancedHTMLFragment @params -Properties "TotalSystem","WithClient","NoClient","WithClient%"

$HTML = $html + $htmlData 

$WorkstationClientCount = $data.WorkstationClientCount.'WithClient%'
$WorkstationNoClientCount = 100 -$data.WorkstationClientCount.'WithClient%'

$html = $html + @" 
        <table width=100%>
        <tr>
          <td style="background-color:$(Set-PercentageColour -Value $WorkstationClientCount);color:#ffffff;" width="$($WorkstationClientCount)%"> $($WorkstationClientCount)% </td>
          <td style="background-color:#eeeeee;color:#333333;" width="$($WorkstationNoClientCount)%"></td>
        </tr>
        </table>

"@


#endregion
#######################################################################
#region HTML All Active Client Heartbeat (DDR) Status
#######################################################################

# Convert results to HTML

                $params = @{'As'='Table';
                    'PreContent'="<h4>&diams; All Active Workstations Client Heartbeat (DDR) Status</h4>";
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $data.ActiveDDRWorkstationCount |
                   ConvertTo-EnhancedHTMLFragment @params -Properties "TotalActive","ActiveHeartBeatDDR","InActiveHeartBeatDDR","ActiveHeartBeatDDR%"

$HTML = $html + $htmlData 

$DDRWorkstatioActive = $data.ActiveDDRWorkstationCount.'ActiveHeartBeatDDR%'
$DDRWorkstatioInactive = 100 -$data.ActiveDDRWorkstationCount.'ActiveHeartBeatDDR%'

$html = $html + @" 
        <table width=100%>
        <tr>
          <td style="background-color:$(Set-PercentageColour -Value $DDRWorkstatioActive);color:#ffffff;" width="$($DDRWorkstatioActive)%"> $($DDRWorkstatioActive)% </td>
          <td style="background-color:#eeeeee;color:#333333;" width="$($DDRWorkstatioInActive)%"></td>
        </tr>
        </table>

"@



#endregion
#######################################################################
#region HTML All Active Client Hardware Inventory Status
#######################################################################

# Convert results to HTML

                $params = @{'As'='Table';
                    'PreContent'="<h4>&diams; All Active Workstations Client Hardware Inventory Status</h4>";
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $data.ActiveHardWareInventoryWorkstationCount |
                   ConvertTo-EnhancedHTMLFragment @params -Properties "TotalActive","ActiveHWInv","InActiveHWInv","ActiveHWInv%"

$HTML = $html + $htmlData 

$WorkstationHWInvActive = $data.ActiveHardWareInventoryWorkstationCount.'ActiveHWInv%'
$WorkstationHWInvInactive = 100 -$data.ActiveHardWareInventoryWorkstationCount.'ActiveHWInv%'

$html = $html + @" 
        <table width=100%>
        <tr>
          <td style="background-color:$(Set-PercentageColour -Value $WorkstationHWInvActive);color:#ffffff;" width="$($WorkstationDDRActive)%"> $($WorkstationDDRActive)% </td>
          <td style="background-color:#eeeeee;color:#333333;" width="$($WorkstationHWInvInactive)%"></td>
        </tr>
        </table>

"@

#endregion
#######################################################################
#region HTML All Active Client Software Inventory Status
#######################################################################

# Convert results to HTML

                $params = @{'As'='Table';
                    'PreContent'="<h4>&diams; All Active Workstations Client Software Inventory Status</h4>";
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $data.ActiveSoftwareInventoryWorkstationCount |
                   ConvertTo-EnhancedHTMLFragment @params -Properties "TotalActive","ActiveSWInv","InActiveSWInv","ActiveSWInv%"


$HTML = $html + $htmlData 

$WorkstationSWInvActive = $data.ActiveSoftwareInventoryWorkstationCount.'ActiveSWInv%'
$WorkstationSWInvInactive = 100 -$data.ActiveSoftwareInventoryWorkstationCount.'ActiveSWInv%'

$html = $html + @" 
        <table width=100%>
        <tr>
          <td style="background-color:$(Set-PercentageColour -Value $WorkstationSWInvActive);color:#ffffff;" width="$($WorkstationSWInvActive)%"> $($WorkstationSWInvActive)% </td>
          <td style="background-color:#eeeeee;color:#333333;" width="$($WorkstationSWInvInactive)%"></td>
        </tr>
        </table>

"@

#endregion
#######################################################################
#region HTML All Active Health Evaluation Status
#######################################################################

# Convert results to HTML




$htmlData = $data.ActiveWorkstationHealthEvalutionCount | 
    ConvertTo-Html -Property "TotalActive","Active_Pass","Active_Fail","Active_Unknown","Active_Pass%" -Head $Style -Body "<Table width=100%><tr><td><h4>All Active Workstation Health Evaluation Status</h4></td></tr></table>" -CssUri "http://www.w3schools.com/lib/w3.css" | 
    Out-String


$HTML = $html + $htmlData 

$WorkstationHealthEvaluationActive = $data.ActiveWorkstationHealthEvalutionCount.'Active_Pass%'
$WorkstationHealthEvaluationInActive = 100 -$data.ActiveWorkstationHealthEvalutionCount.'Active_Pass%'

$html = $html + @" 
        <table width=100%>
        <tr>
          <td style="background-color:$(Set-PercentageColour -Value $WorkstationHealthEvaluationActive);color:#ffffff;" width="$($WorkstationHealthEvaluationActive)%"> $($WorkstationHealthEvaluationActive)% </td>
          <td style="background-color:#eeeeee;color:#333333;" width="$($WorkstationHealthEvaluationInActive)%"></td>
        </tr>
        </table>

"@
#endregion
#######################################################################
#region HTML All Active Client Policy Request Status
#######################################################################

# Convert results to HTML

                $params = @{'As'='Table';
                    'PreContent'="<h4>&diams; All Active Workstation Client PolicyRequest Status</h4>";
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $data.ActiveWorkstationPolicyRequestCount |
                   ConvertTo-EnhancedHTMLFragment @params -Properties "TotalActive","ActivePolicyRequest","InActivePolicyRequest","ActivePolicyRequest%"

$HTML = $html + $htmlData 

$PolicyRequestWorkstationActive = $data.ActiveWorkstationPolicyRequestCount.'ActivePolicyRequest%'
$PolicyRequestWorkstationInactive = 100 -$data.ActiveWorkstationPolicyRequestCount.'ActivePolicyRequest%'

$html = $html + @" 
        <table width=100%>
        <tr>
          <td style="background-color:$(Set-PercentageColour -Value $PolicyRequestWorkstationActive);color:#ffffff;" width="$($PolicyRequestWorkstationActive)%"> $($PolicyRequestWorkstationActive)% </td>
          <td style="background-color:#eeeeee;color:#333333;" width="$($PolicyRequestWorkstationInactive)%"></td>
        </tr>
        </table>

"@
#endregion
#######################################################################
#region HTML All Clients diskspace
#######################################################################

# Convert results to HTML

                $params = @{'As'='Table';
                    'PreContent'="<h4>&diams; Clients with less than 20 gb C-drive </h4>";
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $data.ClientDiskSpace |
                   ConvertTo-EnhancedHTMLFragment @params -Properties Machine,Username,"TotalSpace (GB)","Freespace (GB)"



$HTML = $html + $htmlData 

#endregion
#######################################################################
#region HTML Application Deployment
#######################################################################

# Convert results to HTML

                $params = @{'As'='Table';
                    'PreContent'="<h4>&diams; Applications Deployed last 30 days </h4>";
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $data.ApplicationDeployment |
                   ConvertTo-EnhancedHTMLFragment @params -Properties ApplicationName,AvailableTime,CollectionName,Purpose,Target,Success,"Success%" 


$HTML = $html + $htmlData 
#endregion
#######################################################################
#region HTML Package Deployment
#######################################################################

   # Convert results to HTML

if ($data.PackageDeployment)
{
                  $params = @{'As'='Table';
                    'PreContent'="<h4>&diams; Package Deployed last 30 days </h4>";
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';}

        $htmlData = $data.PackageDeployment |
                   ConvertTo-EnhancedHTMLFragment @params -Properties "ApplicationName","AvailableTime","CollectionName","Purpose","Target","Success","Success%"

$HTML = $html + $htmlData    
}






#endregion
#######################################################################
#region Close html document...
#######################################################################
$html = $html + $post
#endregion
#########################################################
# Mailsettings
# using module Send-MailKitMessage
#########################################################

#use secure connection if available ([bool], optional)
$UseSecureConnectionIfAvailable=$false

#authentication ([System.Management.Automation.PSCredential], optional)
$Credential=[System.Management.Automation.PSCredential]::new("Username", (ConvertTo-SecureString -String "Password" -AsPlainText -Force))

#SMTP server ([string], required)
$SMTPServer=$smtp

#port ([int], required)
$Port=25

#sender ([MimeKit.MailboxAddress] http://www.mimekit.net/docs/html/T_MimeKit_MailboxAddress.htm, required)
$From=[MimeKit.MailboxAddress]$Emailfrom

#recipient list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, required)
$RecipientList=[MimeKit.InternetAddressList]::new()
$RecipientList.Add([MimeKit.InternetAddress]$EmailTo)


#cc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
#$CCList=[MimeKit.InternetAddressList]::new()
#$CCList.Add([MimeKit.InternetAddress]$EmailToCC)



#bcc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
$BCCList=[MimeKit.InternetAddressList]::new()
$BCCList.Add([MimeKit.InternetAddress]"BCCRecipient1EmailAddress")

# Different subject depending on result of search for patches.

#subject ([string], required)
$Subject=[string]"Error Error - Action needed $(get-date)"    

#text body ([string], optional)
#$TextBody=[string]"TextBody"

#HTML body ([string], optional)
$HTMLBody=[string]$html

#attachment list ([System.Collections.Generic.List[string]], optional)
$AttachmentList=[System.Collections.Generic.List[string]]::new()
$AttachmentList.Add("$PSScriptRoot\logo.png")

# Mailparameters
$Parameters=@{
    "UseSecureConnectionIfAvailable"=$UseSecureConnectionIfAvailable    
    #"Credential"=$Credential
    "SMTPServer"=$SMTPServer
    "Port"=$Port
    "From"=$From
    "RecipientList"=$RecipientList
    #"CCList"=$CCList
    #"BCCList"=$BCCList
    "Subject"=$Subject
    #"TextBody"=$TextBody
    "HTMLBody"=$HTMLBody
    "AttachmentList"=$AttachmentList
}
#########################################################
#send email
#########################################################
Send-MailKitMessage @Parameters

#######################################################################
# Test, enable this row to generate html-page
#######################################################################

$html | Out-File -FilePath C:\Temp\test2.html