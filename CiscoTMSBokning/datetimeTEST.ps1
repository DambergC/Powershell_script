param
(
  [datetime]
  [Parameter(Mandatory)]
  $InputStart,
  [string]
  [Parameter(Mandatory)]
  $InputLength
  )


  $starttimeUTC = $InputStart.ToString('yyyy-MM-dd HH:mm:ssZ')

  $Meetingtime = $inputstart.AddHours($InputLength)

  $endtimeUTC = $Meetingtime.ToString('yyyy-MM-dd HH:mm:ssZ')

  $starttimeUTC

  $endtimeUTC



  


