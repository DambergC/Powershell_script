<#
.SYNOPSIS
This script configures multiple Maintenance Windows for a collection.The schedule is based on offset-settings with Patch Tuesday as base.
    
.DESCRIPTION
This script give you options to delete existing Maintance Windows on collection, decide if the Maintance Windows should be for Any installation, Task sequence
or only SoftwareUpdates.

You can also decide which month the Maintance Windows should be configured for.

Some of the funcionality has been borrowed from Daniel Engberg´s script, created 2018 which he borrowed som functionality from Octavian Cordos' script, created in 2015.

####################
Christian Damberg
www.damberg.org
Version 1.0
2021-12-22
####################
    
.EXAMPLE
.\Set-MaintanceWindows.ps1 -CollID ps100137 -OffSetWeeks 1 -OffSetDays 5 -AddStartHour 18 -AddStartMinutes 0 -AddEndHour 4 -AddEndMinutes 0 -PatchMonth "1","2","3","4","5","6","7","8","9","10","11" -patchyear 2022 -ClearOldMW Yes -ApplyTo SoftWareUpdatesOnly
Will create a Maintenance Window with Patch Tuesday + 1 week and 5 days for collection with ID PS100137 for every month except december in 2022. The script also delete old Maintance Windows and the new Maintance Windows are only for SoftwareUpdates.
    
.DISCLAIMER
All scripts and other Powershell references are offered AS IS with no warranty.
These script and functions are tested in my environment and it is recommended that you test these scripts in a test environment before using in your production environment.
#>


PARAM(
    [int]$OffSetWeeks,
    [int]$OffSetDays,
    [Parameter(Mandatory=$True)]
    [int]$AddStartHour,
    [Parameter(Mandatory=$True)]
    [int]$AddStartMinutes,
    [string[]]$PatchMonth,
    [Parameter(Mandatory=$True)]
    [int]$patchyear,
    [string]$FolderName
    )  

#region Initialize


#endregion

#region functions

#Set Patch Tuesday for a Month 
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
   Write-Output "Patch Tuesday this month is $PatchDay"
 }  
 
#Remove all existing Maintenance Windows for a Collection 
Set-Location $PSScriptRoot
 

#Function for append events to logfile located c:\windows\logs
Function Write-Log
{
    PARAM(
    [String]$Message,
    [int]$Severity,
    [string]$Component
    )
    Set-Location $PSScriptRoot
    $Logpath = "C:\Windows\Logs"
    $TimeZoneBias = Get-CimInstance win32_timezone
    $Date= Get-Date -Format "HH:mm:ss.fff"
    $Date2= Get-Date -Format "MM-dd-yyyy"
        "<![LOG[$Message]LOG]!><time=$([char]34)$Date$($TimeZoneBias.bias)$([char]34) date=$([char]34)$date2$([char]34) component=$([char]34)$Component$([char]34) context=$([char]34)$([char]34) type=$([char]34)$Severity$([char]34) thread=$([char]34)$([char]34) file=$([char]34)$([char]34)>"| Out-File -FilePath "$Logpath\Set-MaintenanceWindows.log" -Append -NoClobber -Encoding default

}
# Function to create a folder in Scheduled Task
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


#endregion

#region Parameters

$ErrorMessage = $_.Exception.Message

#endregion



$MonthArray = New-Object System.Globalization.DateTimeFormatInfo 
$MonthNames = $MonthArray.MonthNames 

New-ScheduledTaskFolder $FolderName

#Create Maintance Windows for for every month specified in variable Patchmonth
foreach ($Monthnumber in $PatchMonth) 

{
    #Set Patch Tuesday for each Month 
    $PatchDay = Get-PatchTuesday $Monthnumber $PatchYear
                 
    #Set Maintenance Window Naming Convention (Months array starting from 0 hence the -1) 
    $displaymonth = $Monthnumber - 1 

    #Set Device Collection Maintenace interval  
    $StartTime=$PatchDay.AddDays($OffSetDays).AddHours($AddStartHour).AddMinutes($AddStartMinutes)

    $taskAction = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument '-File C:\scripts\Get-LatestAppLog.ps1'
        
    # Create a new trigger (Daily at 3 AM)
    $tasktrigger = New-ScheduledTaskTrigger -At $StartTime -Once

    # The name of your scheduled task.
    $taskName = "Patchstatus-Mail " +$MonthNames[$displaymonth] + " "+ $patchyear

    # Describe the scheduled task.
    $description = "Export the 10 newest events in the application log"

        Try 
        {
            # Register the scheduled task
            Register-ScheduledTask -TaskName $taskName -Action $taskAction -Trigger $taskTrigger -Description $description -TaskPath $FolderName -User system

            $StartTime
            #Write-Log -Message "Created Maintenance Window $NewMWName for Collection $MWCollection" -Severity 1 -Component "New Maintenance Window"
            #Write-Output "Created Maintenance Window $NewMWName for Collection $MWCollection" 
        }
        Catch 
        {
            Write-Warning "$_.Exception.Message"
            Write-Log -Message "$_.Exception.Message" -Severity 3 -Component "Create new Maintenance Window"
        }

    }

Set-Location $PSScriptRoot
