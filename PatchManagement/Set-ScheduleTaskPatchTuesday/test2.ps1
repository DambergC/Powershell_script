Function New-ScheduledTaskFolder

    {

     Param ($taskpath)

     $ErrorActionPreference = "stop"

     $scheduleObject = New-Object -ComObject schedule.service

     $scheduleObject.connect()

     $rootFolder = $scheduleObject.GetFolder("\")

        Try {$null = $scheduleObject.GetFolder($taskpath)}

        Catch { $null = $rootFolder.CreateFolder($taskpath) }

        Finally { $ErrorActionPreference = "continue" } }

    

Function Create-AndRegisterApplogTask

{

 Param ($taskname, $taskpath)

 $action = New-ScheduledTaskAction -Execute 'Powershell.exe' `

  -Argument '-NoProfile -WindowStyle Hidden -command "& {get-eventlog -logname Application -After ((get-date).AddDays(-1)) | Export-Csv -Path c:\fso\applog.csv -Force -NoTypeInformation}"'

 $trigger =  New-ScheduledTaskTrigger -Daily -At 9am

 Register-ScheduledTask -Action $action -Trigger $trigger -TaskName `

  $taskname -Description "Daily dump of Applog" -TaskPath $taskpath

}

 

Function Create-NewApplotTaskSettings

{

 Param ($taskname, $taskpath)

 $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `

    -Hidden -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -RestartCount 3

 Set-ScheduledTask -TaskName $taskname -Settings $settings -TaskPath $taskpath

}

### ENTRY POINT ###

$taskname = "applog"

$taskpath = "PoshTasks"

If(Get-ScheduledTask -TaskName $taskname -EA 0)

  {Unregister-ScheduledTask -TaskName $taskname -Confirm:$false}

New-ScheduledTaskFolder -taskname $taskname -taskpath $taskpath

Create-AndRegisterApplogTask -taskname $taskname -taskpath $taskpath | Out-Null

Create-NewApplotTaskSettings -taskname $taskname -taskpath $taskpath | Out-Null







$scheduleObject = New-Object -ComObject schedule.service

$scheduleObject.connect()

$rootFolder = $scheduleObject.GetFolder("\")

$rootFolder.DeleteFolder("test",$null)