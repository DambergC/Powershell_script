#
#.Synopsis
#  List all updates member of updategroup in x-numbers of days
#.DESCRIPTION
#   Lists all assigned software updates in a configuration manager 2012 software update group that is selected 
#   from the list of available update groups or provided as a command line option
#.EXAMPLE
#   Get-UpdateGroupMember -script configname

param ($config)


# Path to configfile for script 
$absPath = Join-Path $PSScriptRoot "/$config.XML" 
# Get content of Configfile 
[xml]$config = Get-Content -Path $absPath



$LimitFromXML = $config.scriptsettings.LimitPasttime
$SiteCodeFromXML = $config.scriptsettings.sitecode
$UpdateGroupNameFromXML = $config.scriptsettings.UpdateGroupName
$Emailfrom = $config.scriptsettings.Emailfrom
$EmailTo = $config.scriptsettings.EmailTo
$EmailToCC = $config.scriptsettings.EmailToCC

#Calculate the numbers of days from todays date
$limit = (get-date).AddDays($LimitFromXML)

# Get the powershell module for MEMCM
if (-not(Get-Module -name ConfigurationManager)) {
    Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')
}

# To run the script you must be on ps-drive for MEMCM
Push-Location

Set-Location $SiteCodeFromXML

# Array to collect result
$Result = @()

# Gather all updates in updategrpup
$updates = Get-CMSoftwareUpdate -Fast -UpdateGroupName $UpdateGroupNameFromXML

Write-host "Processing Software Update Group" $UpdateGroupNameFromXML


forEach ($item in $updates)
{
$object = New-Object -TypeName PSObject
$object | Add-Member -MemberType NoteProperty -Name ArticleID -Value $item.ArticleID
$object | Add-Member -MemberType NoteProperty -Name BulletinID -Value $item.BulletinID
$object | Add-Member -MemberType NoteProperty -Name Title -Value $item.LocalizedDisplayName
$object | Add-Member -MemberType NoteProperty -Name LocalizedDescription -Value $item.LocalizedDescription
$object | Add-Member -MemberType NoteProperty -Name DatePosted -Value $item.Dateposted
$object | Add-Member -MemberType NoteProperty -Name Deployed -Value $item.IsDeployed
$object | Add-Member -MemberType NoteProperty -Name 'URL' -Value $item.LocalizedInformativeURL
$result += $object
}

$Title = "Total assigned software updates in " + $UpdateGroupNameFromXML + " = " + $result.count

$UpdatesFound = $result | Sort-Object dateposted -Descending | where-object { $_.dateposted -ge $limit }

# CSS HTML
$header = @"
<style>

    th {

        font-family: Arial, Helvetica, sans-serif;
        color: White;
        font-size: 12px;
        border: 1px solid black;
        padding: 3px;
        background-color: Black;

    } 
    p {

        font-family: Arial, Helvetica, sans-serif;
        color: black;
        font-size: 11px;

    } 
    tr {

        font-family: Arial, Helvetica, sans-serif;
        color: black;
        font-size: 11px;
        vertical-align: text-top;

    } 

    body {
        background-color: #e6f7d0;
      }
      table {
        border: 1px solid black;
        border-collapse: collapse;
      }

      td {
        border: 1px solid black;
        padding: 5px;
        background-color: white;
      }

</style>
"@


if ($UpdatesFound -eq $null )
{
    write-host "No updates downloaded or deployed since $limit"

    $UpdatesFound = @"
    <B>No updates downloaded or deployed since $limit</B><br><br>
    <p></p>
<p>Action needed from third-line support</p>
<ol>Check MEMCM-server</ol>
<ol>Check WSUS</ol>
<ol>Check ...</ol>
"@
}

else 

{

# Text added to mail before list of patches
$pre = @"
<p>The following patches has been downloaded from Microsoft and added to <b><i>$updategroupname</i></b> since $limit</p>
"@

# Text added to mail last 
$post = "<p>Report generated on $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b></p>"

# Mail with pre and post converted to Variable later used to send with send-mailkitmessage
$UpdatesFound = $result | Sort-Object dateposted -Descending | where-object { $_.dateposted -ge $limit }| ConvertTo-Html -Title "Downloaded patches" -PreContent $pre -PostContent $post -Head $header

}


## Mailsettings
# using module Send-MailKitMessage

#use secure connection if available ([bool], optional)
$UseSecureConnectionIfAvailable=$false

#authentication ([System.Management.Automation.PSCredential], optional)
$Credential=[System.Management.Automation.PSCredential]::new("Username", (ConvertTo-SecureString -String "Password" -AsPlainText -Force))

#SMTP server ([string], required)
$SMTPServer="th-ex02.korsberga.local"

#port ([int], required)
$Port=25

#sender ([MimeKit.MailboxAddress] http://www.mimekit.net/docs/html/T_MimeKit_MailboxAddress.htm, required)
$From=[MimeKit.MailboxAddress]$Emailfrom

#recipient list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, required)
$RecipientList=[MimeKit.InternetAddressList]::new()
$RecipientList.Add([MimeKit.InternetAddress]$EmailTo)

#cc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
$CCList=[MimeKit.InternetAddressList]::new()
$CCList.Add([MimeKit.InternetAddress]$EmailToCC)

#bcc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
$BCCList=[MimeKit.InternetAddressList]::new()
$BCCList.Add([MimeKit.InternetAddress]"BCCRecipient1EmailAddress")

# Different subject depending on result of search for patches.
if ($UpdatesFound -eq $null )
{
#subject ([string], required)
$Subject=[string]"Status download patch from Microsoft $(get-date)"
}
else 
{
#subject ([string], required)
$Subject=[string]"Error Error - Action needed $(get-date)"    
}

#text body ([string], optional)
$TextBody=[string]"TextBody"

#HTML body ([string], optional)
$HTMLBody=[string]$UpdatesFound

#attachment list ([System.Collections.Generic.List[string]], optional)
$AttachmentList=[System.Collections.Generic.List[string]]::new()
$AttachmentList.Add("Attachment1FilePath")

# Mailparameters
$Parameters=@{
    "UseSecureConnectionIfAvailable"=$UseSecureConnectionIfAvailable    
    #"Credential"=$Credential
    "SMTPServer"=$SMTPServer
    "Port"=$Port
    "From"=$From
    "RecipientList"=$RecipientList
    "CCList"=$CCList
    #"BCCList"=$BCCList
    "Subject"=$Subject
    #"TextBody"=$TextBody
    "HTMLBody"=$HTMLBody
    #"AttachmentList"=$AttachmentList
}


#send message
Send-MailKitMessage @Parameters

