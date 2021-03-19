
$ConferenceNumber = '2391571'

$pathAPI ='https://tms.cygateviscom.se/tms/external/booking/bookingservice.asmx?wsdl'

$username = 'tms-api.test'

# Read passwordfile och convert to string
$encrypted = Get-Content C:\GitHub\Powershell_script\DV_Cisco_TMS\Keys\cryptfile.txt | ConvertTo-SecureString


# Create variable with username and password
$credential = New-Object System.Management.Automation.PsCredential($username, $encrypted)



$xmlpathConferenceByID = 'C:\GitHub\Powershell_script\DV_Cisco_TMS\GetConferenceByID.XML'

# Import of request-xml to update with new ClientSessionID
$xml=New-Object XML
$xml.Load($xmlpathConferenceByID)
$node=$xml.Envelope.Body.GetConferenceByID
$node.ConferenceId=$ConferenceNumber
$xml.Save($xmlpathConferenceByID)


# POST a new request for an Conference
$PostRequestConferenceByID = (Invoke-WebRequest -Uri $pathAPI -InFile $xmlpathConferenceByID -ContentType 'text/xml' -Method POST -Credential $credential -skiphttpErrorcheck)

$PostRequestConferenceByID.RawContent | Out-File C:\GitHub\Powershell_script\DV_Cisco_TMS\html.xml
$RowInFile = '14'
$callinnumber = Get-Content .\html.xml | Select-Object -Index $RowInFile

$CallinnumberTrimmed = $callinnumber.Trim(" ","-")
$digits = '6'
$callinnumberFinal = $CallinnumberTrimmed.Substring(0,$digits)
$callinnumber
$domain
$CallinnumberTrimmed
$callinnumberFinal