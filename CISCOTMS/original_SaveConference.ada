POST /tms/external/booking/bookingservice.asmx HTTP/1.1
Host: tms.cygateviscom.se
Content-Type: application/soap+xml; charset=utf-8
Content-Length: length

<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Header>
    <ContextHeader xmlns="http://tandberg.net/2004/02/tms/external/booking/">
      <SendConfirmationMail>boolean</SendConfirmationMail>
      <ExcludeConferenceInformation>boolean</ExcludeConferenceInformation>
      <ClientLanguage>string</ClientLanguage>
    </ContextHeader>
    <ExternalAPIVersionSoapHeader xmlns="http://tandberg.net/2004/02/tms/external/booking/">
      <ClientVersionIn>int</ClientVersionIn>
      <ClientIdentifierIn>string</ClientIdentifierIn>
      <ClientLatestNamespaceIn>string</ClientLatestNamespaceIn>
      <NewServiceURL>string</NewServiceURL>
      <ClientSession>string</ClientSession>
    </ExternalAPIVersionSoapHeader>
  </soap12:Header>
  <soap12:Body>
    <SaveConference xmlns="http://tandberg.net/2004/02/tms/external/booking/">
      <Conference>
        <ConferenceId>int</ConferenceId>
        <Title>string</Title>
        <StartTimeUTC>string</StartTimeUTC>
        <EndTimeUTC>string</EndTimeUTC>
        <RecurrenceInstanceIdUTC>string</RecurrenceInstanceIdUTC>
        <RecurrenceInstanceType>string</RecurrenceInstanceType>
        <FirstOccurrenceRecInstanceIdUTC>string</FirstOccurrenceRecInstanceIdUTC>
        <RecurrencePattern>
          <FrequencyType>Daily or DailyWeekday or Weekly or Monthly or Yearly or Secondly or Minutely or Hourly or Default</FrequencyType>
          <Interval>int</Interval>
          <DaysOfWeek>
            <DayOfWeek>Sunday or Monday or Tuesday or Wednesday or Thursday or Friday or Saturday</DayOfWeek>
            <DayOfWeek>Sunday or Monday or Tuesday or Wednesday or Thursday or Friday or Saturday</DayOfWeek>
          </DaysOfWeek>
          <FirstDayOfWeek>Sunday or Monday or Tuesday or Wednesday or Thursday or Friday or Saturday</FirstDayOfWeek>
          <BySetPosition>int</BySetPosition>
          <ByMonthDay>integer</ByMonthDay>
          <PatternEndType>EndByDate or EndByInstances or EndNever or Default</PatternEndType>
          <PatternEndDateUTC>string</PatternEndDateUTC>
          <FirstOccurrenceRecInstanceIdUTC>string</FirstOccurrenceRecInstanceIdUTC>
          <PatternInstances>int</PatternInstances>
          <Exceptions>
            <RecurrenceException xsi:nil="true" />
            <RecurrenceException xsi:nil="true" />
          </Exceptions>
        </RecurrencePattern>
        <OwnerId>long</OwnerId>
        <OwnerUserName>string</OwnerUserName>
        <OwnerFirstName>string</OwnerFirstName>
        <OwnerLastName>string</OwnerLastName>
        <OwnerEmailAddress>string</OwnerEmailAddress>
        <ConferenceType>Reservation Only or Automatic Call Launch or Manual Call Launch or Default or Ad-Hoc conference or One Button To Push or No Connect</ConferenceType>
        <Bandwidth>1b/64kbps or 2b/128kbps or 3b/192kbps or 4b/256kbps or 5b/320kbps or 6b/384kbps or 8b/512kbps or 12b/768kbps or 18b/1152kbps or 23b/1472kbps or 30b/1920kbps or 32b/2048kbps or 48b/3072kbps or 64b/4096kbps or 7b/448kbps or 40b/2560kbps or 96b/6144kbps or Max or 6000kbps or Default</Bandwidth>
        <PictureMode>Continuous Presence or Enhanced CP or Voice Switched or Default</PictureMode>
        <Encrypted>Yes or No or If Possible or Default</Encrypted>
        <DataConference>Yes or No or If Possible or Default</DataConference>
        <ShowExtendOption>Yes or No or Default or AutomaticBestEffort</ShowExtendOption>
        <Password>string</Password>
        <BillingCode>string</BillingCode>
        <ISDNRestrict>boolean</ISDNRestrict>
        <ExternalConference>
          <WebEx>
            <MeetingKey>string</MeetingKey>
            <SipUrl>string</SipUrl>
            <ElementsToExclude>None or MeetingPassword or HostKey or LocalCallInTollFreeNumber or GlobalCallInNumberUrl</ElementsToExclude>
            <MeetingPassword>string</MeetingPassword>
            <JoinMeetingUrl>string</JoinMeetingUrl>
            <HostMeetingUrl>string</HostMeetingUrl>
            <HostKey>string</HostKey>
            <JoinBeforeHostTime>string</JoinBeforeHostTime>
            <Telephony xsi:nil="true" />
            <TmsShouldUpdateMeeting>boolean</TmsShouldUpdateMeeting>
            <SiteUrl>string</SiteUrl>
            <UsePstn>boolean</UsePstn>
            <OwnedExternally>boolean</OwnedExternally>
            <Warnings xsi:nil="true" />
            <Errors xsi:nil="true" />
          </WebEx>
          <ExternallyHosted>
            <DialString>string</DialString>
          </ExternallyHosted>
          <WebExState>
            <WebExInstanceType>Normal or Delete or Modify</WebExInstanceType>
          </WebExState>
        </ExternalConference>
        <EmailTo>string</EmailTo>
        <ConfBundleId>string</ConfBundleId>
        <ConfOwnerId>string</ConfOwnerId>
        <ConferenceInfoText>string</ConferenceInfoText>
        <ConferenceInfoHtml>string</ConferenceInfoHtml>
        <UserMessageText>string</UserMessageText>
        <ExternalSourceId>string</ExternalSourceId>
        <ExternalPrimaryKey>string</ExternalPrimaryKey>
        <DetachedFromExternalSourceId>string</DetachedFromExternalSourceId>
        <DetachedFromExternalPrimaryKey>string</DetachedFromExternalPrimaryKey>
        <Participants>
          <Participant>
            <ParticipantId>int</ParticipantId>
            <NameOrNumber>string</NameOrNumber>
            <EmailAddress>string</EmailAddress>
            <ParticipantCallType>TMS or IP Video <- or IP Tel <- or ISDN Video <- or Telephone <- or IP Video -> or IP Tel -> or ISDN Video -> or Telephone -> or Directory or User or SIP <- or SIP -> or 3G <- or 3G -> or TMS Master Participant or SIP Tel <- or SIP Tel -></ParticipantCallType>
          </Participant>
          <Participant>
            <ParticipantId>int</ParticipantId>
            <NameOrNumber>string</NameOrNumber>
            <EmailAddress>string</EmailAddress>
            <ParticipantCallType>TMS or IP Video <- or IP Tel <- or ISDN Video <- or Telephone <- or IP Video -> or IP Tel -> or ISDN Video -> or Telephone -> or Directory or User or SIP <- or SIP -> or 3G <- or 3G -> or TMS Master Participant or SIP Tel <- or SIP Tel -></ParticipantCallType>
          </Participant>
        </Participants>
        <RecordedConferenceUri>string</RecordedConferenceUri>
        <WebConferencePresenterUri>string</WebConferencePresenterUri>
        <WebConferenceAttendeeUri>string</WebConferenceAttendeeUri>
        <ISDNBandwidth>
          <Bandwidth>1b/64kbps or 2b/128kbps or 3b/192kbps or 4b/256kbps or 5b/320kbps or 6b/384kbps or 8b/512kbps or 12b/768kbps or 18b/1152kbps or 23b/1472kbps or 30b/1920kbps or 32b/2048kbps or 48b/3072kbps or 64b/4096kbps or 7b/448kbps or 40b/2560kbps or 96b/6144kbps or Max or 6000kbps or Default</Bandwidth>
        </ISDNBandwidth>
        <IPBandwidth>
          <Bandwidth>1b/64kbps or 2b/128kbps or 3b/192kbps or 4b/256kbps or 5b/320kbps or 6b/384kbps or 8b/512kbps or 12b/768kbps or 18b/1152kbps or 23b/1472kbps or 30b/1920kbps or 32b/2048kbps or 48b/3072kbps or 64b/4096kbps or 7b/448kbps or 40b/2560kbps or 96b/6144kbps or Max or 6000kbps or Default</Bandwidth>
        </IPBandwidth>
        <ConferenceLanguage>string</ConferenceLanguage>
        <ConferenceTimeZoneRules>
          <TimeZoneRule>
            <ValidFrom>dateTime</ValidFrom>
            <Id>string</Id>
            <DisplayName>string</DisplayName>
            <BaseOffsetInMinutes>int</BaseOffsetInMinutes>
            <Daylight xsi:nil="true" />
            <DaylightOffsetInMinutes>int</DaylightOffsetInMinutes>
            <Standard xsi:nil="true" />
          </TimeZoneRule>
          <TimeZoneRule>
            <ValidFrom>dateTime</ValidFrom>
            <Id>string</Id>
            <DisplayName>string</DisplayName>
            <BaseOffsetInMinutes>int</BaseOffsetInMinutes>
            <Daylight xsi:nil="true" />
            <DaylightOffsetInMinutes>int</DaylightOffsetInMinutes>
            <Standard xsi:nil="true" />
          </TimeZoneRule>
        </ConferenceTimeZoneRules>
        <ConferenceState>
          <Status>All or AllExceptDeleted or Pending or Ongoing or Finished or PendingAndOngoing or MeetingRequest or Rejected or NotSaved or Defective or Deleted</Status>
        </ConferenceState>
        <Version>integer</Version>
        <Location>string</Location>
        <Invitees>string</Invitees>
        <SecretCode>string</SecretCode>
        <MeetingSource>string</MeetingSource>
        <PrivacyFlag>string</PrivacyFlag>
        <JoinUri>string</JoinUri>
        <JoinUrl>string</JoinUrl>
        <TransactionId>long</TransactionId>
      </Conference>
    </SaveConference>
  </soap12:Body>
</soap12:Envelope>