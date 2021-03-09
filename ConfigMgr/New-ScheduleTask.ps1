# Script to create a scheduled task with action, settings, name, description, who will run.

# Create a new task action
$taskAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-ExecutionPolicy Bypass -Noninteractive -File "\\vntsql0081.kvv.se\clienthealth$\ConfigMgrClientHealth.ps1" -Config "\\vntsql0081.kvv.se\clienthealth$\Config.xml" -Webservice "http://vntsql0081.kvv.se/ConfigMgrClientHealth"'

# Triggers
$Tasktrigger = @(
    $(New-ScheduledTaskTrigger -Daily -At 11AM),
    $(New-ScheduledTaskTrigger -AtLogOn))

# Settings
$tasksetting = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 1)

# The name of the scheduled task
$taskName = "ClientHealth"

# Description of the scheduled task
$description = "Clienthealth-script connected to MECM"

# The account that will run the script
$taskPrincipal = New-ScheduledTaskPrincipal -UserId 'NT Instans\SYSTEM' -RunLevel Highest

# Register the scheduled task
Register-ScheduledTask -TaskName $taskName -Action $taskAction -Trigger $taskTrigger -Description $description -Principal $taskPrincipal -Settings $tasksetting

# Get all scheduled task on the client
$taskinstalled = Get-ScheduledTask 

# Verify that the scheduled task is installed
if ($taskinstalled.taskname -eq $taskName) {
    $true
}
else {
    $false
}