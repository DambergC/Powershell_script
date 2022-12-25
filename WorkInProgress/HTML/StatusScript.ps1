
#######################################################################
#region Parameters
#######################################################################

$absPath = Join-Path $PSScriptRoot "/Config.XML"

[xml]$config = Get-Content -Path $absPath

$absPathQuery = Join-Path $PSScriptRoot "/Query.XML"

[xml]$Query = Get-Content -Path $absPathQuery

$finalpath = $config.Settings.Html.Savepath

# Database info
$script:dataSource = $config.Settings.SQLserver.Name # SQL Server name (and instance where applicable)
$script:database = $config.Settings.SQLserver.Database # ConfigMgr Database name


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

# POSHTML5
if (-not(Get-Module -name POSHTML5)) {
    Install-Module -Name POSHTML5 -Verbose
    Import-Module POSHTML5 -Verbose
  }

# PSWriteHTML
if (-not(Get-Module -name PSWriteHTML)) {
    Install-Module -Name PSWriteHTML -Verbose
    Import-Module PSWriteHTML -Verbose
  }  

# PSParseHTML
if (-not(Get-Module -name PSParseHTML)) {
    Install-Module -Name PSParseHTML -Verbose -Force -AllowClobber
    Import-Module PSParseHTML -Verbose -Force    
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


$Sitestatus = Get-SQLData -Query $query.Settings.SiteStatus.Query
$ADRstatus = Get-SQLData -Query $query.Settings.ADRStatus.Query
$MaintanceTaskStatus = Get-SQLData -Query $query.Settings.MaintanceTaskStatus.Query
$SQLDiskStatus = Get-SQLData -Query $query.Settings.SQLDiskStatus.Query



$Process = Get-Process | Select-Object -First 15 | Select-Object name, Priorityclass, fileversion, handles, cpu

Dashboard -Name 'Dashimo Test' -FilePath $PSScriptRoot\DashboardEasy05.html -Show {
    Section -Name 'SiteStatus' {
        Container -Width 600 {
            Panel {
                Table -DataTable $Sitestatus -DisableResponsiveTable {
                    TableButtonPDF
                    TableButtonCopy
                    TableButtonExcel
                    TableButtonPageLength
                } -Buttons @() -DisableSearch -PagingOptions @(5, 10)  -HideFooter
            }

        }
    }
}








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

$html | Out-File -FilePath $finalpath\SiteStatus.html

Set-Location $PSScriptRoot
