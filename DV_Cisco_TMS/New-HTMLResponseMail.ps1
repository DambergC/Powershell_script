﻿$SupportDeskNumber = 'BT 1144-20'
$datum = '2021-04-04'
$StartTime = '09:00'
$EndTime = '10:00'
$PinCode ='1234'
$anslutandesystem = '2'
$ConferenceNumber = '88888888'
$PhoneNumber = '01011223344'
$PhoneNumberINT = '+461011223344'

# Email params
$EmailParams = @{
    To         = 'christian.damberg@cygate.se'
    From       = 'no-reply@cygate.se'
    Smtpserver = 'smtp.cygate.se'
    Subject    = "Bokningsbekräftelse Bryggbokning ärendenr: $SupportDeskNumber  |  $(Get-Date -Format dd-MMM-yyyy)"
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
    Anslutande system:<b>$anslutandesystem</b></p>

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
          <li>Ring:<b><a href=mailto:"$ConferenceNumber@dom.se">$ConferenceNumber@dom.se</a></b></li>
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
<li>Call:<b><a href=mailto:"$ConferenceNumber@dom.se">$ConferenceNumber@dom.se</a></b></li>
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
# Send-MailMessage @EmailParams -Body $html -BodyAsHtml

$html | Out-File C:\dv\test.html -Force

# Send email
#Send-MailMessage @EmailParams -Body $html -BodyAsHtml -Encoding utf8