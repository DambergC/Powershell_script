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
$script:dataSource = 'vax-vs051' # SQL Server name (and instance where applicable)
$script:database = 'CM_sod' # ConfigMgr Database name

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
$sitecode = 'sod:'
$siteserver = 'vax-vs051.sodra.com'
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

.grid { width:100% }

.enhancedhtml-dynamic-table { width:100% }

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
#region HTML Overall Site Status
#######################################################################

$params = @{'As'='Table';
'PreContent'='<h4>Automatic Deployment Rules</h4>';
'EvenRowCssClass'='even';
'OddRowCssClass'='odd';
'MakeTableDynamic'=$true;
'TableCssClass'='grid';}

$html_ADR_Status = $data.ADRStatus |
ConvertTo-EnhancedHTMLFragment @params -Properties name,Status,AutodeploymentEnabled,Lastruntime,Lasterrorcode,Lastrun


#endregion
#######################################################################
#region Close html document...
#######################################################################

$params = @{'CssStyleSheet'=$style;
'Title'="Report for MEMCM Sitecode:$Sitecode SiteServer:$siteserver";
'PreContent'="<h1>Report for MEMCM Sitecode:$Sitecode SiteServer:$siteserver</h1>";
'HTMLFragments'=@($HTMLTop,$HTMLhead,$html_ADR_Status,$HTMLpost)}
#ConvertTo-EnhancedHTML @params | Out-File -FilePath C:\Temp\test2.html
$html = ConvertTo-EnhancedHTML @params
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
#########################################################
#send email
#########################################################
#send-MailKitMessage @Parameters

#######################################################################
# Test, enable this row to generate html-page
#######################################################################

$html | Out-File -FilePath C:\Temp\ADRStatus.html
