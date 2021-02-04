<#
	.SYNOPSIS
		A brief description of the Get-Conference.ps1 file.
	
	.DESCRIPTION
		This script creats an xml for sending a request to CISCO TNS
		
		#https://tms.cygateviscom.se/tms/external/booking/bookingservice.asmx
		
		$url = 'https://tms.cygateviscom.se/tms/external/booking/bookingservice.asmx?wsdl'
		
		$result = (Invoke-WebRequest -Uri $URI -InFile .\sample2.xml -ContentType 'text/xml' -Method POST)
	
	.PARAMETER Conferenceid
		The Conference you want to get info about
	
	.NOTES
		===========================================================================
		Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2021 v5.8.183
		Created on:   	1/27/2021 10:35 AM
		Updated on:		1/27/2021 10:35 AM
		Created by:   	Christian Damberg
		Organization: 	Cygate AB
		Filename:     	Get-Conference.ps1
		===========================================================================
#>
param
(
	[string]$Conferenceid
)



#Variable HEADER
#$conferenceid = '-1' # -1 skapar nytt möte
$SendConfirmationMail = 'Boolean'
$ExcludeConferenceInformation = 'string'
$ClientLanguage = 'string'
$ClientVersionIn = 'int'
$ClientIdentifierIn = 'string'
$ClientLatestNamespaceIn = 'String'
$NewServiceURL = 'string'


$xmlfile = @()

##########################################################################################
# Path for XML output
$XMLpath = 'C:\GitHub\Powershell_script\CISCOTMS\RequestDefaultConference.xml'

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
$writer
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

#Start GetConferenceById
$writer.WriteStartElement("GetConferenceById")
$writer.WriteAttributeString("xmlns", "http://tandberg.net/2004/02/tms/external/booking/")

#ConferenceID
$writer.WriteStartElement("ConferenceID")
$writer.WriteString($conferenceid)
$writer.WriteEndElement()

#end GetConferenceById
$writer.WriteEndElement()
##########################################################################################

#end body
##########################################################################################
$writer.WriteEndElement()
$writer.Flush()
$writer.Close()

[xml]$xmlfile = get-content -Path $XMLpath



#Get-Content .\post_header.xml, .\REQUEST.xml | Set-Content .\Final_RequestFile.xml
