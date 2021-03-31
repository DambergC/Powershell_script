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
	[Parameter(Mandatory = $true, HelpMessage = 'yyyy-MM-dd hh:mm:ss')]
	[datetime]$Inputstart,
	[Parameter(Mandatory = $true, HelpMessage = 'Meeting length in hours')]
	[string]$Inputlength = '12',
	[Parameter(Mandatory = $true)]
	[string]$Bookingnumber,
	[Parameter(Mandatory = $true, HelpMessage = 'the one who booked the conference')]
	[string]$BookedBy
)



$scriptversion = '1.5'
$scriptdate = '2021-03-31'


################################################################################################
# Functions in script
################################################################################################

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

write-log -Level INFO -Message "Script version:$scriptversion ScriptDate:$scriptdate"

# Write to log
write-log -Level INFO -Message "################################################################################################"
write-log -Level INFO -Message "Script version:$scriptversion ScriptDate:$scriptdate"
write-log -Level INFO -Message "New Conference Booking $stamp"
write-log -Level INFO -Message "################################################################################################"
write-log -Level INFO -Message "Path to config: $abspath"
write-log -Level INFO -Message "psscriptroot: $psscriptroot"

################################################################################################
#
# Create Temp-folder if not exist
#
################################################################################################

$Tempfolder = Join-Path $PSScriptRoot $config.ConfigTMS.Tempfolder


if (-not (Test-Path -path $Tempfolder -pathtype Container)) {
    
    try {
        New-Item -Path $Tempfolder -ItemType Directory -ErrorAction Stop | Out-Null #-Force
    }
    catch {
        Write-Error -Message "Unable to create directory '$Tempfolder'. Error was: $_" -ErrorAction Stop
    }
    write-log -level INFO -Message "Successfully created tempfolder $Tempfolder"

}
else {
  write-log -level INFO -Message  "Directory already exist $Tempfolder"
}

################################################################################################
#
# Import user account password from crypted file
#
################################################################################################

# User with rights to run API on Cisco TMS
$username = $config.ConfigTMS.username

# Write to log
write-log -Level INFO -Message "Account running the script: $username"

# Import password from crypted passwordfile
$url = Join-Path $PSScriptRoot $config.ConfigTMS.pathPwdFile

# Write to log
write-log -Level INFO -Message "path to passwordfile: $url"

# Read passwordfile och convert to string
$encrypted = Get-Content $url | ConvertTo-SecureString
$UnsecurePassword = (New-Object PSCredential "user",$encrypted).GetNetworkCredential().Password

# Create variable with username and password
$credential = New-Object System.Management.Automation.PsCredential($username, $encrypted)

################################################################################################
#
# Send Request to get default conference values
#
################################################################################################

# XML file used to extract default-values from Cisco TMS
[System.Xml.XmlDocument] $original_GetDefaultConferenceXML =
@"
<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Header>
    <ExternalAPIVersionSoapHeader xmlns="http://tandberg.net/2004/02/tms/external/booking/">
      <ClientVersionIn>15</ClientVersionIn>
      <ClientIdentifierIn>string</ClientIdentifierIn>
      <ClientLatestNamespaceIn>string</ClientLatestNamespaceIn>
      <NewServiceURL>string</NewServiceURL>
      <ClientSession>string</ClientSession>
    </ExternalAPIVersionSoapHeader>
  </soap12:Header>
  <soap12:Body>
    <GetDefaultConference xmlns="http://tandberg.net/2004/02/tms/external/booking/" />
  </soap12:Body>
</soap12:Envelope>
"@

# Post request to get default values of an conference-request.
$apipath = $config.ConfigTMs.pathCiscoTMSAPI

write-log -Level INFO -Message "Address to Cisco APi $apipath"

$PostRequest = Invoke-WebRequest -Uri $config.ConfigTMs.pathCiscoTMSAPI -Body $original_GetDefaultConferenceXML -ContentType 'text/xml' -Method POST -Credential $credential -UseBasicParsing

# Read XML-response of default values to be used when POST a request for a Conference

[xml]$ResultXML_GetDefaultConference = $PostRequest.Content

################################################################################################
################################################################################################
# This section if for create a xml-file used in the end of this section to post a request to 
# the API for Cisco TMS
################################################################################################
################################################################################################


################################################################################################
#
# Create Request.xml with value from default Conference
#
################################################################################################

# Variable HEADER
$conferenceid = $ResultXML_GetDefaultConference.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.ConferenceId
$SendConfirmationMail = 'true'
$ExcludeConferenceInformation = 'false'
$ClientLanguage = 'en'
$ClientVersionIn = '15'
$ClientIdentifierIn = 'string'
$ClientLatestNamespaceIn = 'String'
$NewServiceURL = 'string'

# Variable from input

# Starttime modified with value from config to adjust for timezone
$StartTimeModified = $Inputstart.AddHours($config.configtms.timeadjust)

# Starttime for meeting
$starttimeUTC = $StartTimeModified.ToString('yyyy-MM-dd HH:mm:ssZ')
$starttimeToMail = $inputstart.ToString('yyyy-MM-dd HH:mm')

# How many hours 
$Meetingtime = $StartTimeModified.AddHours($InputLength)
$endtimeToMail = $inputstart.AddHours($Inputlength)

# End of meeting
$endtimeUTC = $Meetingtime.ToString('yyyy-MM-dd HH:mm:ssZ')
$endtimeToMailformat = $endtimeToMail.ToString('yyyy-MM-dd HH:mm')

$Title = $Bookingnumber

# Variable BODY 
$ClientSession = '0'
$OwnerId = $ResultXML_GetDefaultConference.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.OwnerId
$OwnerUserName = $ResultXML_GetDefaultConference.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.OwnerUserName
$OwnerFirstName = $ResultXML_GetDefaultConference.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.OwnerFirstName
$OwnerLastName = $ResultXML_GetDefaultConference.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.OwnerLastName
$OwnerEmailAddress = $ResultXML_GetDefaultConference.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.OwnerEmailAddress
$ConferenceType = $ResultXML_GetDefaultConference.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.ConferenceType
$Bandwidth = $ResultXML_GetDefaultConference.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.Bandwidth
$PictureMode = $ResultXML_GetDefaultConference.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.PictureMode
$Encrypted = $ResultXML_GetDefaultConference.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.Encrypted
$DataConference = $ResultXML_GetDefaultConference.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.DataConference
$Password = $ResultXML_GetDefaultConference.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.Password
$ShowExtendOption = $ResultXML_GetDefaultConference.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.ShowExtendOption
$ISDNRestrict = $ResultXML_GetDefaultConference.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.ISDNRestrict
$NameOrNumber = $config.configtms.ParticipantDefaultName
$ParticipantCallType = $config.configtms.ParticipantCallType

################################################################################################

# XML file used to post for booking of conference
[System.Xml.XmlDocument] $PostConferenceResult =
@"
<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Header>
    <ContextHeader xmlns="http://tandberg.net/2004/02/tms/external/booking/">
      <SendConfirmationMail>$SendConfirmationMail</SendConfirmationMail>
      <ExcludeConferenceInformation>$ExcludeConferenceInformation</ExcludeConferenceInformation>
      <ClientLanguage>$ClientLanguage</ClientLanguage>
    </ContextHeader>
    <ExternalAPIVersionSoapHeader xmlns="http://tandberg.net/2004/02/tms/external/booking/">
      <ClientVersionIn>$ClientVersionIn</ClientVersionIn>
      <ClientIdentifierIn>$ClientIdentifierIn</ClientIdentifierIn>
      <ClientLatestNamespaceIn>$ClientLatestNamespaceIn</ClientLatestNamespaceIn>
      <NewServiceURL>$NewServiceURL</NewServiceURL>
      <ClientSession>$ClientSession</ClientSession>
    </ExternalAPIVersionSoapHeader>
  </soap12:Header>
  <soap12:Body>
    <SaveConference xmlns="http://tandberg.net/2004/02/tms/external/booking/">
      <Conference>
        <ConferenceID>$conferenceid</ConferenceID>
        <Title>$title</Title>
        <StartTimeUTC>$starttimeUTC</StartTimeUTC>
        <EndTimeUTC>$endtimeUTC</EndTimeUTC>
        <OwnerId>$ownerID</OwnerId>
        <OwnerUserName>$OwnerUserName</OwnerUserName>
        <OwnerFirstName>$OwnerFirstName</OwnerFirstName>
        <OwnerLastName>$OwnerLastName</OwnerLastName>
        <OwnerEmailAddress>$OwnerEmailAddress</OwnerEmailAddress>
        <ConferenceType>$ConferenceType</ConferenceType>
        <Bandwidth>$Bandwidth</Bandwidth>
        <PictureMode>$PictureMode</PictureMode>
        <Encrypted>$Encrypted</Encrypted>
        <DataConference>$DataConference</DataConference>
        <ShowExtendOption>$ShowExtendOption</ShowExtendOption>
        <Password>$Password</Password>
        <ISDNRestrict>$ISDNRestrict</ISDNRestrict>
        <Participants>
          <Participant>
            <NameOrNumber>$NameOrNumber</NameOrNumber>
            <ParticipantCallType>$ParticipantCallType</ParticipantCallType>
          </Participant>
        </Participants>
      </Conference>
    </SaveConference>
  </soap12:Body>
</soap12:Envelope>
"@


################################################################################################
#
# POST Conference Request and get Response if Error 500, get new ClienSessionID
#
################################################################################################

    # POST a request for an Conference
    $PostRequestNewConference = Invoke-WebRequest -Uri $config.ConfigTMs.pathCiscoTMSAPI -Body $PostConferenceResult -ContentType 'text/xml' -Method POST -Credential $credential -skiphttpErrorcheck
    # Get statuscode 200=OK 500
    $StatusCode = $PostRequestNewConference.StatusCode

# Write to log
write-log -Level INFO -Message "ClientSessionID OK (200=OK 500=Expired: $statuscode"

# Read response and if Statuscode is 500 catch new ClientSessionID
[xml]$ConferenceResult = $PostRequestNewConference

# If error 500, extract new ClientSessionID and run the post one more time
if ($statuscode -eq '500') 
{
  [xml]$ConferenceResult = $PostRequestNewConference

  $newClientSessionID = $ConferenceResult.Envelope.Body.Fault.detail.clientsessionid.'#text'
 
  Write-Log -Level INFO -Message "New ClientSessionID $newClientSessionID "
  
  # XML file used to post for booking of conference
  [System.Xml.XmlDocument] $PostConferenceResult =
@"
<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Header>
    <ContextHeader xmlns="http://tandberg.net/2004/02/tms/external/booking/">
      <SendConfirmationMail>$SendConfirmationMail</SendConfirmationMail>
      <ExcludeConferenceInformation>$ExcludeConferenceInformation</ExcludeConferenceInformation>
      <ClientLanguage>$ClientLanguage</ClientLanguage>
    </ContextHeader>
    <ExternalAPIVersionSoapHeader xmlns="http://tandberg.net/2004/02/tms/external/booking/">
      <ClientVersionIn>$ClientVersionIn</ClientVersionIn>
      <ClientIdentifierIn>$ClientIdentifierIn</ClientIdentifierIn>
      <ClientLatestNamespaceIn>$ClientLatestNamespaceIn</ClientLatestNamespaceIn>
      <NewServiceURL>$NewServiceURL</NewServiceURL>
      <ClientSession>$newClientSessionID</ClientSession>
    </ExternalAPIVersionSoapHeader>
  </soap12:Header>
  <soap12:Body>
    <SaveConference xmlns="http://tandberg.net/2004/02/tms/external/booking/">
      <Conference>
        <ConferenceID>$conferenceid</ConferenceID>
        <Title>$title</Title>
        <StartTimeUTC>$starttimeUTC</StartTimeUTC>
        <EndTimeUTC>$endtimeUTC</EndTimeUTC>
        <OwnerId>$ownerID</OwnerId>
        <OwnerUserName>$OwnerUserName</OwnerUserName>
        <OwnerFirstName>$OwnerFirstName</OwnerFirstName>
        <OwnerLastName>$OwnerLastName</OwnerLastName>
        <OwnerEmailAddress>$OwnerEmailAddress</OwnerEmailAddress>
        <ConferenceType>$ConferenceType</ConferenceType>
        <Bandwidth>$Bandwidth</Bandwidth>
        <PictureMode>$PictureMode</PictureMode>
        <Encrypted>$Encrypted</Encrypted>
        <DataConference>$DataConference</DataConference>
        <ShowExtendOption>$ShowExtendOption</ShowExtendOption>
        <Password>$Password</Password>
        <ISDNRestrict>$ISDNRestrict</ISDNRestrict>
        <Participants>
          <Participant>
            <NameOrNumber>$NameOrNumber</NameOrNumber>
            <ParticipantCallType>$ParticipantCallType</ParticipantCallType>
          </Participant>
        </Participants>
      </Conference>
    </SaveConference>
  </soap12:Body>
</soap12:Envelope>
"@
  
  # POST a new request for an Conference
  $PostRequestNewConference = Invoke-WebRequest -Uri $config.ConfigTMs.pathCiscoTMSAPI -Body $PostConferenceResult -ContentType 'text/xml' -Method POST -Credential $credential -skiphttpErrorcheck
    
# Catch the conference values after rerun of invoke-webrequest because of error 500
[xml]$ConferenceResult = $PostRequestNewConference

$TMSConferenceID = $ConferenceResult.Envelope.Body.SaveConferenceResponse.SaveConferenceResult.ConferenceId
}

$TMSConferenceID = $ConferenceResult.Envelope.Body.SaveConferenceResponse.SaveConferenceResult.ConferenceId

# ConferenceID in Cisco TMS
Write-Log -Level INFO -Message "TMSConferenceID $TMSConferenceID "
################################################################################################
################################################################################################
# This section if for extracting the number to call to the conference
################################################################################################
################################################################################################

# XML file to post for more info about the conference to extract the number to call for the Conference.
[System.Xml.XmlDocument] $XMLConferenceByID = @"
<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Header>
    <ContextHeader xmlns="http://tandberg.net/2004/02/tms/external/booking/">
      <SendConfirmationMail>$SendConfirmationMail</SendConfirmationMail>
      <ExcludeConferenceInformation>$ExcludeConferenceInformation</ExcludeConferenceInformation>
      <ClientLanguage>$ClientLanguage</ClientLanguage>
    </ContextHeader>
    <ExternalAPIVersionSoapHeader xmlns="http://tandberg.net/2004/02/tms/external/booking/">
      <ClientVersionIn>$ClientVersionIn</ClientVersionIn>
      <ClientIdentifierIn>$ClientIdentifierIn</ClientIdentifierIn>
      <ClientLatestNamespaceIn>$ClientLatestNamespaceIn</ClientLatestNamespaceIn>
      <NewServiceURL>$NewServiceURL</NewServiceURL>
      <ClientSession>string</ClientSession>
    </ExternalAPIVersionSoapHeader>
  </soap12:Header>
  <soap12:Body>
    <GetConferenceById xmlns="http://tandberg.net/2004/02/tms/external/booking/">
      <ConferenceId>$TMSConferenceID</ConferenceId>
    </GetConferenceById>
  </soap12:Body>
</soap12:Envelope>
"@

# POST a new request for an Conference
$PostRequestConferenceByID = Invoke-WebRequest -Uri $config.ConfigTMs.pathCiscoTMSAPI -Body $XMLConferenceByID -ContentType 'text/xml' -Method POST -Credential $credential -skiphttpErrorcheck

# Send the field RawContent to file
$PostRequestConferenceByID.RawContent | Out-File "$Tempfolder\$TMSConferenceID.xml"

# Read variable in Configfile to which row in file to extract
$RowInFile = $config.ConfigTMS.RowInFile 

# Extract the row with the number to a variable
$callinnumber = Get-Content "$Tempfolder\$TMSConferenceID.xml" | Select-Object -Index $RowInFile

# Extract the number

$CallinnumberFinal = $callinnumber -replace "\D+"

write-log -Level INFO -Message "ConferenceCallInNumber: $callinnumberFinal"

Remove-Item "$Tempfolder\$TMSConferenceID.xml" 

################################################################################################
#
# Create the Email to send to the requester.
#
################################################################################################

# Variabelkonverting för att kunna infoga värden i htmlformat.
$emailsubject = $config.configtms.EmailSubject

# The number to dail in
$ConferenceNumber = $CallinnumberFinal

# Pin-code to the meeting
$PinCode = $ConferenceResult.Envelope.Body.SaveConferenceResponse.SaveConferenceResult.Password

write-log -Level INFO -Message "pwd $pincode  "
write-log -Level INFO -Message "Bookby: $bookedby"

# Phonenumber for national participants
$PhoneNumber = $config.configtms.Phonenumber

# Phonenumber for international participants
$PhoneNumberINT = $config.configtms.PhonenumberINT

#$NumbersOfPaticipants = $ExtDeltagare.count
$domain=$config.configtms.Domain
$JoinUrl=$config.configtms.JoinUrl


# Email params
$EmailParams = @{
    To         = $BookedBy
    From       = $config.ConfigTMS.EmailFrom
    Smtpserver = $config.ConfigTMS.EmailSMTP
    Subject    = "$emailsubject $BookingNumber  |  $(Get-Date -Format dd-MMM-yyyy)"
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
    <h4>BRYGGBOKNING:<b>$BookingNumber</b></h4>
    
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
    
    <p>Ärendenummer:<b>$BookingNumber</b></p>
    <p>Datum och tid:$starttimeToMail - $endtimeToMailformat</p>
    <p>(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna</p> 
    <p><i>Mötesrummet öppnar 5 minuter innan bokad tid.</i></p>

<hr style="height:2px" color="black">
    <p>Mötesnummer:<b>$ConferenceNumber</b><br>
    Pin-kod:<b>$PinCode</b><br>


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
<b><a href="$joinurl">$joinurl</a></b>

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

# Send mail with info about the conference to "bookedby"
Send-MailMessage @EmailParams -Body $html -BodyAsHtml -Encoding utf8
write-log -Level INFO -Message "################################################################################################"
