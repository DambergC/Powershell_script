<#	
    .NOTES
    ===========================================================================
    Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.182
    Created on:   	1/12/2021 10:42 AM
    Created by:   	Christian Damberg
    Organization: 	Cygate AB
    Filename:     	New-Conference.ps1
    ===========================================================================
    .DESCRIPTION
    This script creats an xml for sending a request to CISCO TNS
	
	#https://tms.cygateviscom.se/tms/external/booking/bookingservice.asmx	

	$url = 'https://tms.cygateviscom.se/tms/external/booking/bookingservice.asmx?wsdl'
	
	$result = (Invoke-WebRequest -Uri $URI -InFile .\sample2.xml -ContentType 'text/xml' -Method POST)
#>



#Lista på deltagarna 
$ParticipantList = Import-Csv c:\DV\participant.csv

#Variable HEADER
$conferenceid = '-1' # -1 skapar nytt möte
$SendConfirmationMail = 'Boolean'
$ExcludeConferenceInformation = 'string'
$ClientLanguage = 'string'
$ClientVersionIn = 'int'
$ClientIdentifierIn = 'string'
$ClientLatestNamespaceIn = 'String'
$NewServiceURL = 'string'

#variable BODY
$ClientSession = 'string'
$Title = 'string'
$StartTimeUTC = 'string'
$EndTimeUTC = 'string'
$RecurrenceInstanceIdUTC = 'string'
$RecurrenceInstanceType = 'string'
$FirstOccurrenceRecInstanceIdUTC = 'string'
$FrequencyType = 'Daily or DailyWeekday or Weekly or Monthly or Yearly or Secondly or Minutely or Hourly or Default'
$Interval = 'int'

#dayofweek finns det två rader för... multipla värden?
$DayOfWeek = 'Sunday or Monday or Tuesday or Wednesday or Thursday or Friday or Saturday'

$FirstDayOfWeek = 'Sunday or Monday or Tuesday or Wednesday or Thursday or Friday or Saturday'
$BySetPosition = 'int'
$ByMonthDay = 'integer'
$PatternEndType = 'EndByDate or EndByInstances or EndNever or Default'
$PatternEndDateUTC = 'String'
$FirstOccurrenceRecInstanceIdUTC = 'string'
$PatternInstances = 'int'
$OwnerId = 'long'
$OwnerUserName = 'string'
$OwnerFirstName = 'string'
$OwnerLastName = 'string'
$OwnerEmailAddress = 'string'
$ConferenceType = 'Reservation Only or Automatic Call Launch or Manual Call Launch or Default or Ad-Hoc conference or One Button To Push or No Connect'
$Bandwidth = '1b/64kbps or 2b/128kbps or 3b/192kbps or 4b/256kbps or 5b/320kbps or 6b/384kbps or 8b/512kbps or 12b/768kbps or 18b/1152kbps or 23b/1472kbps or 30b/1920kbps or 32b/2048kbps or 48b/3072kbps or 64b/4096kbps or 7b/448kbps or 40b/2560kbps or 96b/6144kbps or Max or 6000kbps or Default'
$PictureMode = 'Continuous Presence or Enhanced CP or Voice Switched or Default'
$Encrypted = 'Yes or No or If Possible or Default'
$DataConference = 'Yes or No or If Possible or Default'
$DataConference = 'Yes or No or Default or AutomaticBestEffort'
$Password = 'string'
$BillingCode = 'string'
$ISDNRestrict = 'boolean'
$MeetingKey = 'string'
$SipUrl = 'string'
$ElementsToExclude = 'None or MeetingPassword or HostKey or LocalCallInTollFreeNumber or GlobalCallInNumberUrl'
$MeetingPassword = 'string'
$JoinMeetingUrl = 'string'
$HostMeetingUrl = 'string'
$HostKey = 'string'
$JoinBeforeHostTime = 'string'
$TmsShouldUpdateMeeting = 'boolean'
$SiteUrl = 'string'
$UsePstn = 'boolean'
$OwnedExternally = 'boolean'
$DialString = 'string'
$WebExInstanceType = 'Normal or Delete or Modify'
$EmailTo = 'string'
$ConfBundleId = 'string'
$ConfOwnerId = 'string'
$ConferenceInfoText = 'string'
$ConferenceInfoHtml = 'string'
$UserMessageText = 'string'
$ExternalSourceId = 'string'
$ExternalPrimaryKey = 'string'
$DetachedFromExternalSourceId = 'string'
$DetachedFromExternalPrimaryKey = 'string'

#Deltagare - kan vara flera - loop
$ParticipantId = 'int'
$NameOrNumber = 'string'
$EmailAddress = 'string'
$ParticipantCallType = 'TMS or IP Video <- or IP Tel <- or ISDN Video <- or Telephone <- or IP Video -> or IP Tel -> or ISDN Video -> or Telephone -> or Directory or User or SIP <- or SIP -> or 3G <- or 3G -> or TMS Master Participant or SIP Tel <- or SIP Tel ->'
#deltagare

$RecordedConferenceUri = 'string'
$WebConferencePresenterUri = 'string'
$WebConferenceAttendeeUri = 'string'
$ISDNBandwidth = '1b/64kbps or 2b/128kbps or 3b/192kbps or 4b/256kbps or 5b/320kbps or 6b/384kbps or 8b/512kbps or 12b/768kbps or 18b/1152kbps or 23b/1472kbps or 30b/1920kbps or 32b/2048kbps or 48b/3072kbps or 64b/4096kbps or 7b/448kbps or 40b/2560kbps or 96b/6144kbps or Max or 6000kbps or Default'
$ipbandwidth = '1b/64kbps or 2b/128kbps or 3b/192kbps or 4b/256kbps or 5b/320kbps or 6b/384kbps or 8b/512kbps or 12b/768kbps or 18b/1152kbps or 23b/1472kbps or 30b/1920kbps or 32b/2048kbps or 48b/3072kbps or 64b/4096kbps or 7b/448kbps or 40b/2560kbps or 96b/6144kbps or Max or 6000kbps or Default'
$ConferenceLanguage = 'string'

#timezonerule - kan vara flera - loop
$ValidFrom = 'datetime'
$Id = 'string'
$DisplayName = 'string'
$BaseOffsetInMinutes = 'int'
$DaylightOffsetInMinutes = 'int'
#timezonerule

$DaylightOffsetInMinutes = 'All or AllExceptDeleted or Pending or Ongoing or Finished or PendingAndOngoing or MeetingRequest or Rejected or NotSaved or Defective or Deleted'
$Version = 'integer'
$Location = 'string'
$Invitees = 'string'
$SecretCode = 'string'
$MeetingSource = 'string'
$PrivacyFlag = 'string'
$JoinUri = 'string'
$JoinUrl = 'string'
$IsRecordingEnabled = 'boolean'
$TransactionId = 'long'


##########################################################################################
# Path for XML output
$XMLpath = 'c:\DV\REQUEST.xml'
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
$writer.WriteStartElement("SaveConferenc")
$writer.WriteAttributeString("xmlns", "http://tandberg.net/2004/02/tms/external/booking/")
	
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

#RecurrenceInstanceIdUTC
$writer.WriteStartElement("RecurrenceInstanceIdUTC")
$writer.WriteString("$RecurrenceInstanceIdUTC")
$writer.WriteEndElement()

#RecurrenceInstanceType
$writer.WriteStartElement("RecurrenceInstanceType")
$writer.WriteString("$RecurrenceInstanceType")
$writer.WriteEndElement()

#FirstOccurrenceRecInstanceIdUTC
$writer.WriteStartElement("FirstOccurrenceRecInstanceIdUTC")
$writer.WriteString("$FirstOccurrenceRecInstanceIdUTC")
$writer.WriteEndElement()

##########################################################################################
#Start RecurrencePattern
$writer.WriteStartElement("RecurrencePattern")

#FrequencyType
$writer.WriteStartElement("FrequencyType")
$writer.WriteString("$FrequencyType")
$writer.WriteEndElement()

#Interval
$writer.WriteStartElement("Interval")
$writer.WriteString("$Interval")
$writer.WriteEndElement()

##########################################################################################
#Start DaysOfWeek
$writer.WriteStartElement("DaysOfWeek")

#DayOfWeek
$writer.WriteStartElement("DayOfWeek")
$writer.WriteString("$DayOfWeek")
$writer.WriteEndElement()

#DayOfWeek
$writer.WriteStartElement("DayOfWeek")
$writer.WriteString("$DayOfWeek")
$writer.WriteEndElement()

#end DaysOfWeek
$writer.WriteEndElement()
##########################################################################################

#FirstDayOfWeek
$writer.WriteStartElement("FirstDayOfWeek")
$writer.WriteString("$FirstDayOfWeek")
$writer.WriteEndElement()

#BySetPosition
$writer.WriteStartElement("BySetPosition")
$writer.WriteString("$BySetPosition")
$writer.WriteEndElement()

#ByMonthDay
$writer.WriteStartElement("ByMonthDay")
$writer.WriteString("$ByMonthDay")
$writer.WriteEndElement()

#PatternEndType
$writer.WriteStartElement("PatternEndType")
$writer.WriteString("$PatternEndType")
$writer.WriteEndElement()

#PatternEndDateUTC
$writer.WriteStartElement("PatternEndDateUTC")
$writer.WriteString("$PatternEndDateUTC")
$writer.WriteEndElement()

#FirstOccurrenceRecInstanceIdUTC
$writer.WriteStartElement("FirstOccurrenceRecInstanceIdUTC")
$writer.WriteString("$FirstOccurrenceRecInstanceIdUTC")
$writer.WriteEndElement()

#PatternInstances
$writer.WriteStartElement("PatternInstances")
$writer.WriteString("$PatternInstances")
$writer.WriteEndElement()

##########################################################################################
#Start Exceptions
$writer.WriteStartElement("Exceptions")

#RecurrenceException
$writer.WriteStartElement("RecurrenceException")
$writer.WriteAttributeString("xsi:nil", "true")
$writer.WriteEndElement()

#RecurrenceException
$writer.WriteStartElement("RecurrenceException")
$writer.WriteAttributeString("xsi:nil", "true")
$writer.WriteEndElement()

#end Exceptions
$writer.WriteEndElement()
##########################################################################################

#end RecurrencePattern
$writer.WriteEndElement()
##########################################################################################

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

#DataConference
$writer.WriteStartElement("DataConference")
$writer.WriteString("$DataConference")
$writer.WriteEndElement()

#Password
$writer.WriteStartElement("Password")
$writer.WriteString("$Password")
$writer.WriteEndElement()

#BillingCode
$writer.WriteStartElement("BillingCode")
$writer.WriteString("$BillingCode")
$writer.WriteEndElement()

#ISDNRestrict
$writer.WriteStartElement("ISDNRestrict")
$writer.WriteString("$ISDNRestrict")
$writer.WriteEndElement()

##########################################################################################
#Start ExternalConference
$writer.WriteStartElement("ExternalConference")

##########################################################################################
#Start WebEx
$writer.WriteStartElement("WebEx")

#MeetingKey
$writer.WriteStartElement("MeetingKey")
$writer.WriteString("$MeetingKey")
$writer.WriteEndElement()

#SipUrl
$writer.WriteStartElement("SipUrl")
$writer.WriteString("$SipUrl")
$writer.WriteEndElement()

#ElementsToExclude
$writer.WriteStartElement("ElementsToExclude")
$writer.WriteString("$ElementsToExclude")
$writer.WriteEndElement()

#MeetingPassword
$writer.WriteStartElement("MeetingPassword")
$writer.WriteString("$MeetingPassword")
$writer.WriteEndElement()

#JoinMeetingUrl
$writer.WriteStartElement("JoinMeetingUrl")
$writer.WriteString("$JoinMeetingUrl")
$writer.WriteEndElement()

#HostMeetingUrl
$writer.WriteStartElement("HostMeetingUrl")
$writer.WriteString("$HostMeetingUrl")
$writer.WriteEndElement()

#HostKey
$writer.WriteStartElement("HostKey")
$writer.WriteString("$HostKey")
$writer.WriteEndElement()

#JoinBeforeHostTime
$writer.WriteStartElement("JoinBeforeHostTime")
$writer.WriteString("$JoinBeforeHostTime")
$writer.WriteEndElement()

#Telephony
$writer.WriteStartElement("Telephony")
$writer.WriteAttributeString("xsi:nil", "true")
$writer.WriteEndElement()

#TmsShouldUpdateMeeting
$writer.WriteStartElement("TmsShouldUpdateMeeting")
$writer.WriteString("$TmsShouldUpdateMeeting")
$writer.WriteEndElement()

#SiteUrl
$writer.WriteStartElement("SiteUrl")
$writer.WriteString("$SiteUrl")
$writer.WriteEndElement()

#UsePstn
$writer.WriteStartElement("UsePstn")
$writer.WriteString("$UsePstn")
$writer.WriteEndElement()

#OwnedExternally
$writer.WriteStartElement("OwnedExternally")
$writer.WriteString("$OwnedExternally")
$writer.WriteEndElement()

#Warnings
$writer.WriteStartElement("Warnings")
$writer.WriteAttributeString("xsi:nil", "true")
$writer.WriteEndElement()

#Errors
$writer.WriteStartElement("Errors")
$writer.WriteAttributeString("xsi:nil", "true")
$writer.WriteEndElement()

#end WebEx
$writer.WriteEndElement()
##########################################################################################

##########################################################################################
#Start ExternallyHosted
$writer.WriteStartElement("ExternallyHosted")

#DialString
$writer.WriteStartElement("DialString")
$writer.WriteString("$DialString")
$writer.WriteEndElement()

#end ExternallyHosted
$writer.WriteEndElement()
##########################################################################################

##########################################################################################
#Start WebExState
$writer.WriteStartElement("WebExState")

#WebExInstanceType
$writer.WriteStartElement("WebExInstanceType")
$writer.WriteString("$WebExInstanceType")
$writer.WriteEndElement()

#end WebExState
$writer.WriteEndElement()
##########################################################################################

#end ExternalConference
$writer.WriteEndElement()
##########################################################################################

#EmailTo
$writer.WriteStartElement("EmailTo")
$writer.WriteString("$EmailTo")
$writer.WriteEndElement()

#ConfBundleId
$writer.WriteStartElement("ConfBundleId")
$writer.WriteString("$ConfBundleId")
$writer.WriteEndElement()

#ConfOwnerId
$writer.WriteStartElement("ConfOwnerId")
$writer.WriteString("$ConfOwnerId")
$writer.WriteEndElement()

#ConferenceInfoText
$writer.WriteStartElement("ConferenceInfoText")
$writer.WriteString("$ConferenceInfoText")
$writer.WriteEndElement()

#ConferenceInfoHtml
$writer.WriteStartElement("ConferenceInfoHtml")
$writer.WriteString("$ConferenceInfoHtml")
$writer.WriteEndElement()

#UserMessageText
$writer.WriteStartElement("UserMessageText")
$writer.WriteString("$UserMessageText")
$writer.WriteEndElement()

#ExternalSourceId
$writer.WriteStartElement("ExternalSourceId")
$writer.WriteString("$ExternalSourceId")
$writer.WriteEndElement()

#ExternalPrimaryKey
$writer.WriteStartElement("ExternalPrimaryKey")
$writer.WriteString("$ExternalPrimaryKey")
$writer.WriteEndElement()

#DetachedFromExternalSourceId
$writer.WriteStartElement("DetachedFromExternalSourceId")
$writer.WriteString("$DetachedFromExternalSourceId")
$writer.WriteEndElement()

#DetachedFromExternalPrimaryKey
$writer.WriteStartElement("DetachedFromExternalPrimaryKey")
$writer.WriteString("$DetachedFromExternalPrimaryKey")
$writer.WriteEndElement()

##########################################################################################
#Start Participants
$writer.WriteStartElement("Participants")

##########################################################################################
##########################################################################################

$ParticipantList.foreach(
	{
	$NameOrNumber = $($_.namn)
	$EmailAddress = $($_.epost)
	$ParticipantCallType = 'IP Video'

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

#RecordedConferenceUri
$writer.WriteStartElement("RecordedConferenceUri")
$writer.WriteString("$RecordedConferenceUri")
$writer.WriteEndElement()

#WebConferencePresenterUri
$writer.WriteStartElement("WebConferencePresenterUri")
$writer.WriteString("$WebConferencePresenterUri")
$writer.WriteEndElement()

#WebConferenceAttendeeUri
$writer.WriteStartElement("WebConferenceAttendeeUri")
$writer.WriteString("$WebConferenceAttendeeUri")
$writer.WriteEndElement()

##########################################################################################
#Start ISDNBandwidth
$writer.WriteStartElement("ISDNBandwidth")

#Bandwidth
$writer.WriteStartElement("Bandwidth")
$writer.WriteString("$ISDNBandwidth")
$writer.WriteEndElement()

#end ISDNBandwidth
$writer.WriteEndElement()
##########################################################################################

##########################################################################################
#Start IPBandwidth
$writer.WriteStartElement("IPBandwidth")

#Bandwidth
$writer.WriteStartElement("Bandwidth")
$writer.WriteString("$ipbandwidth")
$writer.WriteEndElement()

#end IPBandwidth
$writer.WriteEndElement()
##########################################################################################

#ConferenceLanguage
$writer.WriteStartElement("ConferenceLanguage")
$writer.WriteString("$ConferenceLanguage")
$writer.WriteEndElement()

##########################################################################################
##########################################################################################
#Start ConferenceTimeZoneRules
$writer.WriteStartElement("ConferenceTimeZoneRules")

##########################################################################################
#Start TimeZoneRule
$writer.WriteStartElement("TimeZoneRule")

#ValidFrom
$writer.WriteStartElement("ValidFrom")
$writer.WriteString("$ValidFrom")
$writer.WriteEndElement()

#Id
$writer.WriteStartElement("Id")
$writer.WriteString("$Id")
$writer.WriteEndElement()

#DisplayName
$writer.WriteStartElement("DisplayName")
$writer.WriteString("$DisplayName")
$writer.WriteEndElement()

#BaseOffsetInMinutes
$writer.WriteStartElement("BaseOffsetInMinutes")
$writer.WriteString("$BaseOffsetInMinutes")
$writer.WriteEndElement()

#Daylight
$writer.WriteStartElement("Daylight")
$writer.WriteAttributeString("xsi:nil", "true")
$writer.WriteEndElement()

#DaylightOffsetInMinutes
$writer.WriteStartElement("DaylightOffsetInMinutes")
$writer.WriteString("$DaylightOffsetInMinutes")
$writer.WriteEndElement()

#Standard
$writer.WriteStartElement("Standard")
$writer.WriteAttributeString("xsi:nil", "true")
$writer.WriteEndElement()

#end TimeZoneRule
$writer.WriteEndElement()
##########################################################################################

#end ConferenceTimeZoneRules
$writer.WriteEndElement()
##########################################################################################
##########################################################################################

##########################################################################################
#Start ConferenceState
$writer.WriteStartElement("ConferenceState")

#DaylightOffsetInMinutes
$writer.WriteStartElement("DaylightOffsetInMinutes")
$writer.WriteString("$DaylightOffsetInMinutes")
$writer.WriteEndElement()

#end ConferenceState
$writer.WriteEndElement()
##########################################################################################

#Version
$writer.WriteStartElement("Version")
$writer.WriteString("$Version")
$writer.WriteEndElement()

#Location
$writer.WriteStartElement("Location")
$writer.WriteString("$Location")
$writer.WriteEndElement()

#Invitees
$writer.WriteStartElement("Invitees")
$writer.WriteString("$Invitees")
$writer.WriteEndElement()

#SecretCode
$writer.WriteStartElement("SecretCode")
$writer.WriteString("$SecretCode")
$writer.WriteEndElement()

#MeetingSource
$writer.WriteStartElement("MeetingSource")
$writer.WriteString("$MeetingSource")
$writer.WriteEndElement()

#PrivacyFlag
$writer.WriteStartElement("PrivacyFlag")
$writer.WriteString("$PrivacyFlag")
$writer.WriteEndElement()

#JoinUri
$writer.WriteStartElement("JoinUri")
$writer.WriteString("$JoinUri")
$writer.WriteEndElement()

#JoinUrl
$writer.WriteStartElement("JoinUrl")
$writer.WriteString("$JoinUrl")
$writer.WriteEndElement()

#IsRecordingEnabled
$writer.WriteStartElement("IsRecordingEnabled")
$writer.WriteString("$IsRecordingEnabled")
$writer.WriteEndElement()

#TransactionId
$writer.WriteStartElement("TransactionId")
$writer.WriteString("$TransactionId")
$writer.WriteEndElement()

#end ConferenceResult
$writer.WriteEndElement()
##########################################################################################

#end SaveConferenceResponse
$writer.WriteEndElement()
##########################################################################################

#end body
##########################################################################################
$writer.WriteEndElement()
$writer.Flush()
$writer.Close()

Get-Content .\post_header.xml, .\REQUEST.xml | Set-Content Final_RequestFile.xml
