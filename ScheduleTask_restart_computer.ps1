$ParamsTrigger = @{
    Once = $true
    At   = Get-Date "2022-09-30 09:50:00"
}

#remove old task
Unregister-ScheduledTask -TaskName "Once Reboot" -Confirm:$false -ErrorAction SilentlyContinue

# Create task action
$taskAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument 'Restart-Computer -Force'
# Create a trigger (Mondays at 4 AM)
$taskTrigger = New-ScheduledTaskTrigger @ParamsTrigger
# The user to run the task
$taskUser = New-ScheduledTaskPrincipal -UserId "LOCALSERVICE" -LogonType ServiceAccount -RunLevel Highest
# The name of the scheduled task.
$taskName = "Once Reboot"
# Describe the scheduled task.
$description = "Forcibly reboot the computerat specific time"
# Register the scheduled task
Register-ScheduledTask -TaskName $taskName -Action $taskAction -Trigger $taskTrigger -Principal $taskUser -Description $description
