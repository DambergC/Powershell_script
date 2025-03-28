﻿<#
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

$absPath = Join-Path $PSScriptRoot "/Config.XML"

[xml]$config = Get-Content -Path $absPath

$finalpath = $config.Settings.Html.Savepath

# Database info
$script:dataSource = $config.Settings.SQLserver.Name # SQL Server name (and instance where applicable)
$script:database = $config.Settings.SQLserver.Database # ConfigMgr Database name

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
$sitecode = $config.Settings.SiteServer.SiteCode
$siteserver = $config.Settings.SiteServer.Name
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
$emailto = 'christian.damberg@cygate.se'
#
# The email when the script cant find any updates
$email_noErrors = 'christian.damberg@trivselhus.se'
#
# SMTP-server
$smtp = 'webmail.trivselhus.se'

#endregion

#Call myScript1 from myScript2
invoke-expression -Command .\Get-MEMCMStatus_ADR.ps1

#######################################################################
#region Functions
#######################################################################

# Check for ConfigMgr 
if (-not(Get-Module -name ConfigurationManager)) {
  Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')
}

# MAilkitMessage
if (-not(Get-Module -name send-mailkitmessage)) {
  Install-Module send-mailkitmessage -Verbose -Force
  Import-Module send-mailkitmessage -Verbose -Force
}

# EnhancedHTML2
if (-not(Get-Module -name EnhancedHTML2)) {
  Install-Module -Name EnhancedHTML2 -Verbose
  Import-Module EnhancedHTML2 -Verbose
}

<# POSHTML5
if (-not(Get-Module -name POSHTML5)) {
    Install-Module -Name POSHTML5 -Verbose
    Import-Module POSHTML5 -Verbose
  }

#>

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
#endregion

#######################################################################
#region Powershell Region
#######################################################################


#######################################################################
#Query Region
#######################################################################
# Create has table to store data
$Data = @{}
#######################################################################
# SYSTEM
#######################################################################

###########################################
#region QUERY
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



#######################################################################
#region Create html header
#######################################################################
# Html CSS style
$HTMLTop = @"
<!doctype html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
"@

$style = @"
<style>
/* Remove space around the email design. */
html,
body {
    margin: 0 auto !important;
    padding: 0 !important;
    height: 100% !important;
    width: 100% !important;
}

/* Stop Outlook resizing small text. */
* {
    -ms-text-size-adjust: 100%;
}

/* Stop Outlook from adding extra spacing to tables. */
table,
td {
    mso-table-lspace: 0pt !important;
    mso-table-rspace: 0pt !important;
}

/* Use a better rendering method when resizing images in Outlook IE. */
img {
    -ms-interpolation-mode:bicubic;
}

/* Prevent Windows 10 Mail from underlining links. Styles for underlined links should be inline. */
a {
    text-decoration: none;
}

body {
  
    color:#333333;
    font-family:Calibri,Tahoma;
    font-size: 10pt;
    !important; mso-line-height-rule: exactly;
}

h1 {
    text-align:Left;
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

.odd  { background-color:#9bddff; }

.even { background-color:#fffacd; }

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

.grid { width:600px }

.enhancedhtml-dynamic-table { width:600px }

.red {
    color:red;
    font-weight:bold;
} 
</style>
"@

$HTMLhead = @"
</head>
<body width=“100%” style=“margin: 0; padding: 0 !important; mso-line-height-rule: exactly;”>
<br>
<img src='cid:logo.png' height="50">
<br>
<p>This report don´t fix the problem, you have to investigate the problem your self :-)</p>
"@


#########################################################
# Footer of the email
#########################################################
$HTMLpost = @"
<p>Raport created $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b></p>
<p>Script created by:<br><a href="mailto:Your Email">Your name</a><br>
<a href="https://your blog">your description of your blog</a>
</body>
</html>
"@

#endregion

#######################################################################
#region HTML 
#######################################################################

$params = @{'As'='list';
'PreContent'="<h4>Disk & SQL status </h4>";
'EvenRowCssClass'='even';
'OddRowCssClass'='odd';
'MakeTableDynamic'=$false;
'TableCssClass'='grid';}

$html_Disk_SQL_Status = $data.Disksitesql | ConvertTo-EnhancedHTMLFragment @params -Properties status,'site system',"Object Type",Total,Free,%Free

Dashboard -Name 'Dashimo Test' -FilePath $PSScriptRoot\Dashboard.html {
  Tab -Name 'Forest' {
      Section -Name 'Forest Information' -Invisible {
          Section -Name 'Status SQL and Disk' {
              Table -HideFooter -DataTable $data.DiskSiteSQL
          }
          Section -Name 'FSMO Roles' {
              Table -HideFooter -DataTable $data.ADRStatus
          }

      }

  }
}


#endregion
#######################################################################
#region Close html document...
#######################################################################

$params = @{'CssStyleSheet'=$style;
#'Title'="Report for MEMCM Sitecode:$Sitecode SiteServer:$siteserver";
#'PreContent'="<h1>Report for MEMCM Sitecode:$Sitecode SiteServer:$siteserver</h1>";
#'HTMLFragments'=@($HTMLTop,$HTMLhead,$html_Disk_SQL_Status,$HTMLpost)}
'HTMLFragments'=@($html_Disk_SQL_Status)}

$html = ConvertTo-EnhancedHTML @params
#endregion
#########################################################
#region
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
$Subject=[string]"Daily report for MEMCM SiteCode:$sitecode on SiteServer:$siteserver $(get-date)"    

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
#endregion
#########################################################
#send email
#########################################################
#send-MailKitMessage @Parameters

#######################################################################
# Test, enable this row to generate html-page
#######################################################################



$html | Out-File -FilePath $finalpath\DiskSQLStatus.html


