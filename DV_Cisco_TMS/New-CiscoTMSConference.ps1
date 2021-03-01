<#
	.SYNOPSIS
		A brief description of the New-CiscoTMSConference.ps1 file.
	
	.DESCRIPTION
		A detailed description of the New-CiscoTMSConference.ps1 file.
	
	.PARAMETER Inputstart
		A description of the Inputstart parameter.
	
	.PARAMETER Inputlength
		A description of the Inputlength parameter.
	
	.PARAMETER Bookingnumber
		A description of the Bookingnumber parameter.
	
	.PARAMETER BookedBy
		A description of the BookedBy parameter.
	
	.PARAMETER ExtDeltagare
		A description of the ExtDeltagare parameter.
	
	.PARAMETER TestOnly
		Generate XML only
	
	.NOTES
		===========================================================================
		Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.182
		Created on:   	02/04/2021 1:50 PM
		Created by:   	Christian Damberg, Sebastian Thörngren
		Organization: 	Cygate AB
		Filename:     	New-CiscoTMSConference.ps1
		===========================================================================
#>


# Inputvalues needed to send a request for Conference.
[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true,
			   HelpMessage = 'yyyy-MM-dd hh:mm:ss')]
	[datetime]$Inputstart,
	[Parameter(Mandatory = $true,
			   HelpMessage = 'Meeting length in hours')]
	[string]$Inputlength,
	[Parameter(Mandatory = $true)]
	[string]$Bookingnumber,
	[Parameter(Mandatory = $true,
			   HelpMessage = 'the one who booked the conference')]
	[string]$BookedBy,
	[Parameter(Mandatory = $true)]
	[array]$ExtDeltagare,
	[Parameter(Mandatory = $false)]
	[ValidateSet('Yes', 'No', IgnoreCase = $true)]
	[string]$TestOnly = 'Yes'
)

################################################################################################
# Functions in script
################################################################################################

# Function used in troubleshooting during development
function Write-Failure {
$global:helpme = $body
$global:helpmoref = $moref
$global:result = $_.Exception.Response.GetResponseStream()
$global:reader = New-Object System.IO.StreamReader($global:result)
$global:responseBody = $global:reader.ReadToEnd();
Write-Host -BackgroundColor:Black -ForegroundColor:Red "Status: A system exception was caught."
Write-Host -BackgroundColor:Black -ForegroundColor:Red $global:responsebody
Write-Host -BackgroundColor:Black -ForegroundColor:Red "The request body has been saved to `$global:helpme"
break
}

# Function to write to logfile
Function Write-Log {
  [CmdletBinding()]
  Param(
  [Parameter(Mandatory=$False)]
  [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
  [String]
  $Level = "INFO",

  [Parameter(Mandatory=$True)]
  [string]
  $Message
  )

  $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
  $Line = "$Stamp $Level $Message"
        Add-Content $logfile -Value $Line
  }

################################################################################################

# Path to configfile for script
$absPath = Join-Path $PSScriptRoot "/CiscoTMSConfig.XML"

# Get content of Configfile
[xml]$config = Get-Content -Path $absPath

################################################################################################

# Path to logfile
$logfile = Join-Path $PSScriptRoot $config.ConfigTMS.Logfile

# Write to log
write-log -Level INFO -Message "################# ny bokning $stamp ###############"
write-log -Level INFO -Message "Path to config: $abspath"
write-log -Level INFO -Message "psscriptroot: $psscriptroot"


################################################################################################
#
# Import user account password from crypted file
#
################################################################################################

# User with rights to run API on Cisco TMS
$username = $config.ConfigTMS.username

# Write to log
write-log -Level INFO -Message "Username: $username"

# Import password from crypted passwordfile
$url = Join-Path $PSScriptRoot $config.ConfigTMS.pathPwdFile

# Write to log
write-log -Level INFO -Message "path to passwordfile: $url"

# Read passwordfile och convert to string
$encrypted = Get-Content $url | ConvertTo-SecureString

# Create variable with username and password
$credential = New-Object System.Management.Automation.PsCredential($username, $encrypted)

################################################################################################
#
# Send Request to get default conference values
#
################################################################################################

# Create variable to use with invoke-request
$DefaultConferenceXML = Join-Path $PSScriptRoot $config.ConfigTMS.pathDefaultConferenceXML

# Write to log
write-log -Level INFO -Message "path to DefaultConferenceXML $DefaultConferenceXML"

# Post request to get default values of an conference-request.
$PostRequest = (Invoke-WebRequest -Uri $config.ConfigTMs.pathCiscoTMSAPI -InFile $DefaultConferenceXML -ContentType 'text/xml' -Method POST -Credential $credential)


# Read XML-response of default values to be used when POST a request for a Conference
[xml]$DefaultConfValue = $PostRequest

# List Value in console DEVELOPMENT ONLY
if ($TestOnly -eq 'Yes') 
{
  $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult
}

################################################################################################
#
# Create Request.xml with value from default Conference
#
################################################################################################

# Variable HEADER
$conferenceid = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.ConferenceId
$SendConfirmationMail = 'true'
$ExcludeConferenceInformation = 'false'
$ClientLanguage = 'en'
$ClientVersionIn = '15'
$ClientIdentifierIn = 'string'
$ClientLatestNamespaceIn = 'String'
$NewServiceURL = 'string'

# Variable from input
$starttimeUTC = $InputStart.ToString('yyyy-MM-dd HH:mm:ssZ')

# $SwedishTime = 

$Meetingtime = $inputstart.AddHours($InputLength)

$endtimeUTC = $Meetingtime.ToString('yyyy-MM-dd HH:mm:ssZ')

$Title = $Bookingnumber

# Variable BODY 
$ClientSession = '0'
$OwnerId = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.OwnerId
$OwnerUserName = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.OwnerUserName
$OwnerFirstName = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.OwnerFirstName
$OwnerLastName = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.OwnerLastName
$OwnerEmailAddress = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.OwnerEmailAddress
$ConferenceType = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.ConferenceType
$Bandwidth = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.Bandwidth
$PictureMode = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.PictureMode
$Encrypted = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.Encrypted
$DataConference = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.DataConference
$Password = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.Password
$ShowExtendOption = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.ShowExtendOption
$ISDNRestrict = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.ISDNRestrict

################################################################################################
# Path for XML output
$XMLpath = Join-Path $PSScriptRoot $config.ConfigTMS.PathConferenceRequestOutput

# Write to log
write-log -Level INFO -Message "Path to Requestfile $xmlpath"

################################################################################################

################################################################################################
# Set up encoding, and create new instance of XMLTextWriter
$encoding = [System.Text.Encoding]::UTF8
$writer = New-Object -TypeName System.Xml.XmlTextWriter -ArgumentList ($XMLpath, $encoding)
$writer.Formatting = [system.xml.formatting]::indented
$writer.Indentation = 2
################################################################################################

# Write start of XML document - REQUEST
$writer.WriteStartDocument()

# Start envelope
################################################################################################
$writer.WriteStartElement("soap12:Envelope")
$writer.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
$writer.WriteAttributeString("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
$writer.WriteAttributeString("xmlns:soap12", "http://www.w3.org/2003/05/soap-envelope")

# Start header
################################################################################################

# Start soap12:header
$writer.WriteStartElement("soap12:Header")

# Start ContextHeader
$writer.WriteStartElement("ContextHeader")
$writer.WriteAttributeString("xmlns", "http://tandberg.net/2004/02/tms/external/booking/")

# SendConfirmationMail
$writer.WriteStartElement("SendConfirmationMail")
$writer.WriteString("$SendConfirmationMail")
$writer.WriteEndElement()

# ExcludeConferenceInformation
$writer.WriteStartElement("ExcludeConferenceInformation")
$writer.WriteString("$ExcludeConferenceInformation")
$writer.WriteEndElement()

# ClientLanguage
$writer.WriteStartElement("ClientLanguage")
$writer.WriteString("$ClientLanguage")
$writer.WriteEndElement()

# End ContextHeader
$writer.WriteEndElement()

# Start ExternalAPIVersionSoapHeader
$writer.WriteStartElement("ExternalAPIVersionSoapHeader")
$writer.WriteAttributeString("xmlns", "http://tandberg.net/2004/02/tms/external/booking/")

# ClientVersionIn
$writer.WriteStartElement("ClientVersionIn")
$writer.WriteString("$ClientVersionIn")
$writer.WriteEndElement()

# ClientIdentifierIn
$writer.WriteStartElement("ClientIdentifierIn")
$writer.WriteString("$ClientIdentifierIn")
$writer.WriteEndElement()

# ClientLatestNamespaceIn
$writer.WriteStartElement("ClientLatestNamespaceIn")
$writer.WriteString("$ClientLatestNamespaceIn")
$writer.WriteEndElement()

# NewServiceURL
$writer.WriteStartElement("NewServiceURL")
$writer.WriteString("$NewServiceURL")
$writer.WriteEndElement()

# ClientSession
$writer.WriteStartElement("ClientSession")
$writer.WriteString("$ClientSession")
$writer.WriteEndElement()



# End ExternalAPIVersionSoapHeader
$writer.WriteEndElement()

# End soap12:header
$writer.WriteEndElement()

# Start body
################################################################################################
$writer.WriteStartElement("soap12:Body")

# Start SaveConference
$writer.WriteStartElement("SaveConference")
$writer.WriteAttributeString("xmlns", "http://tandberg.net/2004/02/tms/external/booking/")

    # Start Conference
    $writer.WriteStartElement("Conference")

# ConferenceID
$writer.WriteStartElement("ConferenceID")
$writer.WriteString($conferenceid)
$writer.WriteEndElement()

# Title
$writer.WriteStartElement("Title")
$writer.WriteString("$Title")
$writer.WriteEndElement()

# StartTimeUTC
$writer.WriteStartElement("StartTimeUTC")
$writer.WriteString("$StartTimeUTC")
$writer.WriteEndElement()

# EndTimeUTC
$writer.WriteStartElement("EndTimeUTC")
$writer.WriteString("$EndTimeUTC")
$writer.WriteEndElement()

# OwnerId
$writer.WriteStartElement("OwnerId")
$writer.WriteString("$OwnerId")
$writer.WriteEndElement()

# OwnerUserName
$writer.WriteStartElement("OwnerUserName")
$writer.WriteString("$OwnerUserName")
$writer.WriteEndElement()

# OwnerFirstName
$writer.WriteStartElement("OwnerFirstName")
$writer.WriteString("$OwnerFirstName")
$writer.WriteEndElement()

# OwnerLastName
$writer.WriteStartElement("OwnerLastName")
$writer.WriteString("$OwnerLastName")
$writer.WriteEndElement()

# OwnerEmailAddress
$writer.WriteStartElement("OwnerEmailAddress")
$writer.WriteString("$OwnerEmailAddress")
$writer.WriteEndElement()

# ConferenceType
$writer.WriteStartElement("ConferenceType")
$writer.WriteString("$ConferenceType")
$writer.WriteEndElement()

# Bandwidth
$writer.WriteStartElement("Bandwidth")
$writer.WriteString("$Bandwidth")
$writer.WriteEndElement()

# PictureMode
$writer.WriteStartElement("PictureMode")
$writer.WriteString("$PictureMode")
$writer.WriteEndElement()

# Encrypted
$writer.WriteStartElement("Encrypted")
$writer.WriteString("$Encrypted")
$writer.WriteEndElement()

# DataConference
$writer.WriteStartElement("DataConference")
$writer.WriteString("$DataConference")
$writer.WriteEndElement()

# ShowExtendOption
$writer.WriteStartElement("ShowExtendOption")
$writer.WriteString("$ShowExtendOption")
$writer.WriteEndElement()

# Password
$writer.WriteStartElement("Password")
$writer.WriteString("$Password")
$writer.WriteEndElement()

# ISDNRestrict
$writer.WriteStartElement("ISDNRestrict")
$writer.WriteString("$ISDNRestrict")
$writer.WriteEndElement()

################################################################################################
# Start Participants
$writer.WriteStartElement("Participants")

################################################################################################
################################################################################################

$ExtDeltagare.ForEach(
	{
		$NameOrNumber = $_
		$ParticipantCallType = 'IP Video ->'
		
		# Start Participant
		$writer.WriteStartElement("Participant")
		
		# NameOrNumber
		$writer.WriteStartElement("NameOrNumber")
		$writer.WriteString("$NameOrNumber")
		$writer.WriteEndElement()
		
		# ParticipantCallType
		$writer.WriteStartElement("ParticipantCallType")
		$writer.WriteString("$ParticipantCallType")
		$writer.WriteEndElement()
		
		# End Participant
		$writer.WriteEndElement()
	}
)
################################################################################################

# End Participants
$writer.WriteEndElement()
################################################################################################
################################################################################################

    # End Conference
    $writer.WriteEndElement()

# End SaveConference
$writer.WriteEndElement()
################################################################################################

# End body
################################################################################################
$writer.WriteEndElement()
$writer.Flush()
$writer.Close()


################################################################################################
#
# POST Conference Request and get Response
#
################################################################################################



try
{
    # POST a request for an Conference
    $PostRequestNewConference = (Invoke-WebRequest -Uri $config.ConfigTMs.pathCiscoTMSAPI -InFile $XMLpath -ContentType 'text/xml' -Method POST -Credential $credential -skiphttpErrorcheck)
    
    # Get statuscode 200=OK 500
    $StatusCode = $PostRequestNewConference.StatusCode
    
}

catch
{
    $StatusCode = $_.Exception.Response.StatusCode.value__
    $failure = $_.Exception.Response.body
    Write-Failure

}



# Test only writes statuscode in terminal
if ($TestOnly -eq 'yes') {

  write-host "Statuscode:$StatusCode"
  $global:helpme
  $global:responsebody
}


# Write to log
write-log -Level INFO -Message "Statuskod för bokningen 200=OK 500=ClientSession_Expired: $statuscode"


# Read response and if Statuscode is 500 catch new ClientSessionID
[xml]$ConferenceResult = $PostRequestNewConference

if ($statuscode -eq '500') 
{

  # Export Result to XML
  $ConferenceResult | Export-Clixml -Path C:\dv\SEB-V2\ConferenceResult500.xml

  # Import XML and extract new ClientSessionID
  [XML]$indata = Get-Content -Path C:\dv\SEB-V2\ConferenceResult500.xml
  [XML]$raw = $indata.Objs.XD
  $newClientSessionID = $raw.Envelope.Body.Fault.detail.clientsessionid.'#text'

  Write-Log -Level INFO -Message "New ClientSessionID $newClientSessionID "
  
  # Import of request-xml to update with new ClientSessionID
  $xml=New-Object XML
  $xml.Load($XMLpath)
  $node=$xml.Envelope.Header.ExternalAPIVersionSoapHeader
  $node.ClientSession=$newClientSessionID
  $xml.Save($XMLpath)

  # POST a new request for an Conference
  $PostRequestNewConference = (Invoke-WebRequest -Uri $config.ConfigTMs.pathCiscoTMSAPI -InFile $XMLpath -ContentType 'text/xml' -Method POST -Credential $credential -skiphttpErrorcheck)
    
# Write log  

}



################################################################################################
#
# Create the Email to send to the requester.
#
################################################################################################

# Variabelkonverting för att kunna infoga värden i htmlformat.
$emailsubject = $config.configtms.EmailSubject
$ConferenceNumber = $ConferenceResult.Envelope.Body.SaveConferenceResponse.SaveConferenceResult.ConferenceId
$StartTime = $ConferenceResult.Envelope.Body.SaveConferenceResponse.SaveConferenceResult.StartTimeUTC
$EndTime = $ConferenceResult.Envelope.Body.SaveConferenceResponse.SaveConferenceResult.EndTimeUTC
$PinCode = $ConferenceResult.Envelope.Body.SaveConferenceResponse.SaveConferenceResult.Password
$PhoneNumber = $config.configtms.Phonenumber
$PhoneNumberINT = $config.configtms.PhonenumberINT
$SupportDeskNumber = $Bookingnumber
$NumbersOfPaticipants = $ExtDeltagare.count
$domain=$config.configtms.Domain


# Email params
$EmailParams = @{
    To         = $BookedBy
    From       = $config.ConfigTMS.EmailFrom
    Smtpserver = $config.ConfigTMS.EmailSMTP
    Subject    = "$emailsubject $Ordernumber  |  $(Get-Date -Format dd-MMM-yyyy)"
}



# Create html header whit stylesheet
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
    h4
  {
      background-color:Tomato;
      color:white;
      text-align: center;
  }

    p
  {
        font-size: 13px;
  }
    ul
  {
        font-size: 13px;
  }
</style>
"@


# Create html header whit stylesheet
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
    h4
  {
      background-color:Tomato;
      color:white;
      text-align: center;
  }

    p
  {
        font-size: 13px;
  }
    ul
  {
        font-size: 13px;
  }
</style>
"@


# Set html
$html = $html + @"
<table cellpadding="10" cellspacing="10">
<tr>
  <td>
    <h4>BRYGGBOKNING:<b>$SupportDeskNumber</b></h4>
    
    <p><b>Hej!</b></p>
    <p>Här kommer din videokonferensebokning.</p>
    <p>Kontrollera gärna att tidpunkt och datum stämmer.</p>
    <p>Observera att mötesrummet nedan endast finns tillgängligt för den tiden som ni har bokat. Var vänliga och kontakta Tekniksupporten om ni behöver ändra något.</p>
    <p><b><u>Samtliga deltagare ska ringa</u></b> in till bryggans mötesrum. Om ytterligare deltagare behöver vara med via telefon eller videokonferens under tiden som ni är i bryggan, behöver även de ringa in till bryggans mötesrum.</p>
    <p>All info om hur man ringer och vad man bör tänka på vid t ex. skyddad deltagare, står nedan.På <a href="https://intranatet.dom.se/stod-och-verktyg/it-teknik-och-telefoni/videokonferens/webrtc/">intranätet</a> finns information och lathundar (på svenska och engelska) om WebRTC som ni gärna får bifoga till deltagarna.</p>
    <p>Se till att alla får informationen om hur de ska ringa in. Det är alltså <u><b>upp till er att förmedla informationen vidare till berörda parter</b></u> så att de vet hur de ska ringa.</p>
    <p>Det finns instruktioner på engelska längst ner.</p> 
    <p>Då det finns en begränsad mängd platser i videokonferensbryggan var vänlig och kontakta Tekniksupporten <a href="mailto:teknik@dom.se">teknik@dom.se</a> om ni inte längre behöver bokningen eller vill justera bokningen.</p> 
    <p>Vid problem kontaktar personal inom Sveriges Domstolar Tekniksupporten. Externa parter kontaktar först sin videokonferensansvarig för att säkerställa att problemet inte hos dem, därefter kontaktar parten domstolen.</p>
    <p>Med Vänlig Hälsning<br> Tekniksupporten</p> 
    <p>Tfn: 22 000, tonval 3</p> 

<hr style="height:2px" color="black">

    <h4>Du har bokat in följande virtuella mötesrum i Sveriges Domstolars brygga:</h4>
    
    <p><b>$SupportDeskNumber</b></p>
    <p>Datum och tid:$datum,$StartTime - $EndTime</p>
    <p>(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna</p> 
    <p><i>Mötesrummet öppnar 5 minuter innan bokad tid.</i></p>

<hr style="height:2px" color="black">
    <p>Mötesnummer:<b>$ConferenceNumber</b><br>
    Pin-kod:<b>$PinCode</b><br>
    Anslutande system:<b>$NumbersOfPaticipants</b></p>

<hr style="height:2px" color="black">

    <h4>Instruktioner</h4>

      <table style="color:red">
      <tr>
      <h5>Vid skyddad deltagare</h5> 
      <ul>
          <li>Tänk på att inte berätta eller visa var denne sitter i videokonferenssamtalet.</li>
          <li>Du måste även förmedla till organisationen där den skyddade parten sitter att det är en skyddad part och att även de måste vara försiktiga med vad de säger och visar i videokonferenssamtalet.</li>
          <li>Man måste även vara generellt försiktig med hur man lämnar ut information om målet och var den skyddade parten eventuellt sitter.</li>
      </ul>
      </tr>
      </table>

    <p><U><b>Sveriges Domstolar: Videokonferens och JabberVideo-användare</b></u><br>
      <ul>
          <li>Ring:<b>$ConferenceNumber</b>
          <li>PIN-kod från Sal: Välj "Sänd tonval" i pekskärmen. Skicka PIN-kod:<b>$PinCode#</b></li>
          <li>PIN-kod från Rum: Aktivera tonval med knapp # på fjärrkontrollen. Skicka PIN-kod:<b>$PinCode#</b></li>
          <li>PIN-kod från JabberVideo/Movi: Välj "Tonval". Skicka PIN-kod:<b>$PinCode#</b></li>
      </ul>
    </p>

    <p><u><B>Deltagare via Internet/SGSI (utanför Sveriges Domstolar)</b></u><br>
      <ul>
          <li>Ring:<b><a href=mailto:"$ConferenceNumber@$domain">$ConferenceNumber@$domain</a></b></li>
          <li>Med tonval/knappsats slå PIN-kod:<b>$PinCode#</b></li> 
      </ul>
    </p>

<p><u><b>Deltagare via webbläsare - WebRTC (utanför Sveriges Domstolar)</u></b><br>
Bild- och ljudkvalitet kan variera beroende på din dators/plattas/mobiltelefones prestanda och bandbredd.<br> 
Webb-kamera och headset rekommenderas.</p>
<ul>
<li>Videokonferens via webbläsare (fungerar ej i Internet Explorer eller Edge version 41):</li>
<b> Cisco Meeting App link ska in här....!!!!</b>

<li>Mötesnummer:<b>$ConferenceNumber</b></li>
<li>PIN-kod:<b>$PinCode</b></li> 
</ul>


<p><u><b>Telefondeltagare och Videokonferens via ISDN (utanför Sveriges Domstolar)</b></u><br>
<ul>
<li>Ring:<b>$PhoneNumber</b></li>
<li>Med tonval/knappsats slå mötesnummer:<b>$ConferenceNumber#</b></li> 
<li>Med tonval/knappsats slå PIN-kod:<b>$PinCode#</b></li> 
</ul>
<hr style="height:2px" color="black">

<p><B>Booking confirmation: Virtual meeting room in the Swedish Courts MCU</b></p>

<p><b><u>Participant via Internet/SGSI (outside the Swedish National Courts)</u></b></p>
<ul>
<li>Call:<b><a href=mailto:"$ConferenceNumber@$domain">$ConferenceNumber@$domain</a></b></li>
<li>Send Pin:<b>$PinCode#</b></li> 
</ul>
</p>

<p><u><b>Participant via web browser - WebRTC (outside the Swedish National Courts)</u></b><br></p>
<ul>
<li>Join using web browser (not Internet Explorer or Edge version 41):</li>
<b> Cisco Meeting App link ska in här....!!!!</b>

<li>Meeting number:<b>$ConferenceNumber</b></li>
<li>PIN:<b>$PinCode</b></li> 
</ul>

<p><u><b>Participant by phone or Video via ISDN (outside the Swedish National Courts)</b></u><br>
<ul>
<li>Call ISDN to:<b>$PhoneNumberINT</b></li>
<li>Send meeting number:<b>$ConferenceNumber#</b></li> 
<li>Send Pin:<b>$PinCode#</b></li> 
</ul>


  </td>

</tr>
</table>
"@


# Close html document
$html = $html + @"
</body>
</html>
"@

# Send email

$testfile_html = Join-Path $PSScriptRoot $config.ConfigTMS.PathTESTHtmlOutFile

if ($TestOnly -eq 'Yes')
{
	$html | Out-File $testfile_html -Force
}
else 
{
	Send-MailMessage @EmailParams -Body $html -BodyAsHtml -Encoding utf8
}

