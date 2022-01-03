# Create a new task action
$taskAction = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument '-File C:\scripts\Get-LatestAppLog.ps1'
$taskAction

# Create a new trigger (Daily at 3 AM)
$taskTrigger = New-ScheduledTaskTrigger -Daily -At 3PM
$tasktrigger = New-ScheduledTaskTrigger -At "2019-10-01T05:00:00Z" -Once
$tasktrigger

# Register the new PowerShell scheduled task

# The name of your scheduled task.
$taskName = "ExportAppLog"

# Describe the scheduled task.
$description = "Export the 10 newest events in the application log"

# Register the scheduled task
Register-ScheduledTask `
    -TaskName $taskName `
    -Action $taskAction `
    -Trigger $taskTrigger `
    -Description $description

    Function Get-PatchTuesday ($Month,$Year)  
    { 
       $FindNthDay=2 #Aka Second occurence 
       $WeekDay='Tuesday' 
       $todayM=($Month).ToString()
       $todayY=($Year).ToString()
       $StrtMonth=$todayM+'/1/'+$todayY 
       [datetime]$StrtMonth=$todayM+'/1/'+$todayY 
       while ($StrtMonth.DayofWeek -ine $WeekDay ) { $StrtMonth=$StrtMonth.AddDays(1) } 
       $PatchDay=$StrtMonth.AddDays(7*($FindNthDay-1)) 
       return $PatchDay
       Write-Log -Message "Patch Tuesday this month is $PatchDay" -Severity 1 -Component "Set Patch Tuesday"
      <#
      .SYNOPSIS
      Short description
      
      .DESCRIPTION
      Long description
      
      .PARAMETER Month
      Parameter description
      
      .PARAMETER Year
      Parameter description
      
      .EXAMPLE
      An example
      
      .NOTES
      General notes
      #>Write-Output "Patch Tuesday this month is $PatchDay"
    }      