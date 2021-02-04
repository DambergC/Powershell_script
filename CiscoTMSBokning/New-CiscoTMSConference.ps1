
<# Set and encrypt credentials to file using default method #>

#$credential = Get-Credential
#$credential.Password | ConvertFrom-SecureString | Set-Content c:\dv\scriptsencrypted_password1.txt

#####################################################################
#
# Del 1: Läs in behörigheter för Cisco TMS
#
#####################################################################

#Användaren som kör mot TMS
$username = "tms-api.test"

#Inläsning av lösenordet från den krypterade filen
$encrypted = Get-Content c:\dv\scriptsencrypted_password1.txt | ConvertTo-SecureString

#Skapar variabeln för inloggning
$credential = New-Object System.Management.Automation.PsCredential($username, $encrypted)

#sökvägen till API för Cisco TMS
$url = 'https://tms.cygateviscom.se/tms/external/booking/bookingservice.asmx?wsdl'

#####################################################################
#
# Del 2: Skicka Request för defaultvärden för en bokning
#
#####################################################################

#Skicka en förfrågan om defaultvärden från Cisco TMS
$PostRequest = (Invoke-WebRequest -Uri $url -InFile 'C:\dv\original_GetDefaultConference.xml' -ContentType 'text/xml' -Method POST -Credential $credential)

#Läser in värden från förfrågan om defaultvärden
[xml]$DefaultConfValue = $PostRequest

#UTV - Lista värden från default
$DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult



#####################################################################
#
# Del 3: Skapa en Request-xml för en bokning
#
#####################################################################

#Deltagarlista
$ParticipantList = Import-Csv 'C:\dv\participant.csv' -Encoding UTF8

#Variable HEADER
$conferenceid = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.ConferenceId
$SendConfirmationMail = 'true'
$ExcludeConferenceInformation = 'false'
$ClientLanguage = 'en'
$ClientVersionIn = '15'
$ClientIdentifierIn = 'string'
$ClientLatestNamespaceIn = 'String'
$NewServiceURL = 'string'

#variable BODY
#$ClientSession = '87e4dcbc-79fd-4ef7-adaf-b8df6038bcb8'
$ClientSession = 'string'
$Title = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.Title
$StartTimeUTC = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.StartTimeUTC
$EndTimeUTC = $DefaultConfValue.Envelope.Body.GetDefaultConferenceResponse.GetDefaultConferenceResult.EndTimeUTC

#dayofweek finns det tvÃ¥ rader fÃ¶r... multipla vÃ¤rden?

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

##########################################################################################
# Path for XML output
$XMLpath = 'c:\dv\ConfRequest.xml'
##########################################################################################

##########################################################################################
# Set up encoding, and create new instance of XMLTextWriter
$encoding = [System.Text.Encoding]::UTF8
$writer = New-Object -TypeName System.Xml.XmlTextWriter -ArgumentList ($XMLpath, $encoding)
$writer.Formatting = [system.xml.formatting]::indented
$writer.Indentation = 2
##########################################################################################

# Write start of XML document - REQUEST
$writer.WriteStartDocument()

#start envelope
##########################################################################################
$writer.WriteStartElement("soap12:Envelope")
$writer.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
$writer.WriteAttributeString("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
$writer.WriteAttributeString("xmlns:soap12", "http://www.w3.org/2003/05/soap-envelope")

#start header
##########################################################################################

#start soap12:header
$writer.WriteStartElement("soap12:Header")

#start ContextHeader
$writer.WriteStartElement("ContextHeader")
$writer.WriteAttributeString("xmlns", "http://tandberg.net/2004/02/tms/external/booking/")

#SendConfirmationMail
$writer.WriteStartElement("SendConfirmationMail")
$writer.WriteString("$SendConfirmationMail")
$writer.WriteEndElement()

#ExcludeConferenceInformation
$writer.WriteStartElement("ExcludeConferenceInformation")
$writer.WriteString("$ExcludeConferenceInformation")
$writer.WriteEndElement()

#ClientLanguage
$writer.WriteStartElement("ClientLanguage")
$writer.WriteString("$ClientLanguage")
$writer.WriteEndElement()

#end ContextHeader
$writer.WriteEndElement()

#start ExternalAPIVersionSoapHeader
$writer.WriteStartElement("ExternalAPIVersionSoapHeader")
$writer.WriteAttributeString("xmlns", "http://tandberg.net/2004/02/tms/external/booking/")

#ClientVersionIn
$writer.WriteStartElement("ClientVersionIn")
$writer.WriteString("$ClientVersionIn")
$writer.WriteEndElement()

#ClientIdentifierIn
$writer.WriteStartElement("ClientIdentifierIn")
$writer.WriteString("$ClientIdentifierIn")
$writer.WriteEndElement()

#ClientLatestNamespaceIn
$writer.WriteStartElement("ClientLatestNamespaceIn")
$writer.WriteString("$ClientLatestNamespaceIn")
$writer.WriteEndElement()

#NewServiceURL
$writer.WriteStartElement("NewServiceURL")
$writer.WriteString("$NewServiceURL")
$writer.WriteEndElement()

#ClientSession
$writer.WriteStartElement("ClientSession")
$writer.WriteString("$ClientSession")
$writer.WriteEndElement()



#end ExternalAPIVersionSoapHeader
$writer.WriteEndElement()

#end soap12:header
$writer.WriteEndElement()

#start body
##########################################################################################
$writer.WriteStartElement("soap12:Body")

#Start SaveConference
$writer.WriteStartElement("SaveConference")
$writer.WriteAttributeString("xmlns", "http://tandberg.net/2004/02/tms/external/booking/")

    #Start Conference
    $writer.WriteStartElement("Conference")

#ConferenceID
$writer.WriteStartElement("ConferenceID")
$writer.WriteString($conferenceid)
$writer.WriteEndElement()

#Title
$writer.WriteStartElement("Title")
$writer.WriteString("$Title")
$writer.WriteEndElement()

#StartTimeUTC
$writer.WriteStartElement("StartTimeUTC")
$writer.WriteString("$StartTimeUTC")
$writer.WriteEndElement()

#EndTimeUTC
$writer.WriteStartElement("EndTimeUTC")
$writer.WriteString("$EndTimeUTC")
$writer.WriteEndElement()

#OwnerId
$writer.WriteStartElement("OwnerId")
$writer.WriteString("$OwnerId")
$writer.WriteEndElement()

#OwnerUserName
$writer.WriteStartElement("OwnerUserName")
$writer.WriteString("$OwnerUserName")
$writer.WriteEndElement()

#OwnerFirstName
$writer.WriteStartElement("OwnerFirstName")
$writer.WriteString("$OwnerFirstName")
$writer.WriteEndElement()

#OwnerLastName
$writer.WriteStartElement("OwnerLastName")
$writer.WriteString("$OwnerLastName")
$writer.WriteEndElement()

#OwnerEmailAddress
$writer.WriteStartElement("OwnerEmailAddress")
$writer.WriteString("$OwnerEmailAddress")
$writer.WriteEndElement()

#ConferenceType
$writer.WriteStartElement("ConferenceType")
$writer.WriteString("$ConferenceType")
$writer.WriteEndElement()

#Bandwidth
$writer.WriteStartElement("Bandwidth")
$writer.WriteString("$Bandwidth")
$writer.WriteEndElement()

#PictureMode
$writer.WriteStartElement("PictureMode")
$writer.WriteString("$PictureMode")
$writer.WriteEndElement()

#Encrypted
$writer.WriteStartElement("Encrypted")
$writer.WriteString("$Encrypted")
$writer.WriteEndElement()

#DataConference
$writer.WriteStartElement("DataConference")
$writer.WriteString("$DataConference")
$writer.WriteEndElement()

#ShowExtendOption
$writer.WriteStartElement("ShowExtendOption")
$writer.WriteString("$ShowExtendOption")
$writer.WriteEndElement()

#Password
$writer.WriteStartElement("Password")
$writer.WriteString("$Password")
$writer.WriteEndElement()

#ISDNRestrict
$writer.WriteStartElement("ISDNRestrict")
$writer.WriteString("$ISDNRestrict")
$writer.WriteEndElement()

##########################################################################################
#Start Participants
$writer.WriteStartElement("Participants")

##########################################################################################
##########################################################################################

$ParticipantList.ForEach(
  {
    $ParticipantId = $($_.Pnr)
    $EmailAddress = $($_.epost)
    $NameOrNumber = $($_.deltagare)
    $ParticipantCallType = 'IP Video ->'
		
    #Start Participant
    $writer.WriteStartElement("Participant")
		
    #ParticipantId
    $writer.WriteStartElement("ParticipantId")
    $writer.WriteString("$ParticipantId")
    $writer.WriteEndElement()
		
    #NameOrNumber
    $writer.WriteStartElement("NameOrNumber")
    $writer.WriteString("$NameOrNumber")
    $writer.WriteEndElement()
		
    #EmailAddress
    $writer.WriteStartElement("EmailAddress")
    $writer.WriteString("$EmailAddress")
    $writer.WriteEndElement()
		
    #ParticipantCallType
    $writer.WriteStartElement("ParticipantCallType")
    $writer.WriteString("$ParticipantCallType")
    $writer.WriteEndElement()
		
    #end Participant
    $writer.WriteEndElement()
  }
)
##########################################################################################

#end Participants
$writer.WriteEndElement()
##########################################################################################
##########################################################################################

    #end Conference
    $writer.WriteEndElement()

#end SaveConference
$writer.WriteEndElement()
##########################################################################################

#end body
##########################################################################################
$writer.WriteEndElement()
$writer.Flush()
$writer.Close()


#####################################################################
#
# Del 4: Skicka en request för bokning.
#
#####################################################################


$PostRequestNewConference = (Invoke-WebRequest -Uri $url -InFile 'C:\dv\ConfRequest.xml' -ContentType 'text/xml' -Method POST -Credential $credential)

[xml]$ConferenceResult = $PostRequestNewConference