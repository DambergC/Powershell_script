<#
 
************************************************************************************************************************
 
Created:    2015-05-27
Version:    1.2
 
Disclaimer:
This script is provided "AS IS" with no warranties, confers no rights and 
is not supported by the authors or DeploymentArtist.
 
Author - Johan Arwidmark
    Twitter: @jarwidmark
    Blog   : https://deploymentresearch.com
 
************************************************************************************************************************
 
#>
 
 
Import-Module $env:SMS_ADMIN_UI_PATH.Replace("\bin\i386","\bin\configurationmanager.psd1")
$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-Location "$($SiteCode.Name):\"
 
$Collection = "MW-test"
$ADRName = "Client Updates 2015"
$SiteServer = $Env:COMPUTERNAME
$SiteServerFQDN = "$SiteServer.$Env:USERDNSDOMAIN"
$DeployPackageLocation = "\\$SiteServer\repository\Software Updates\$ADRName"
 
$Products = "Windows 7","Office 2010"
$UpdateClassifications = "Critical Updates","Security Updates","Service Packs","Update Rollups","Updates"
$Severity = "Critical","Important","Moderate"
 
if (Get-CMDeviceCollection -Name $Collection)
{
    # All good
    Write-Output "$Collection collection found, continue"
}
Else
{
    Write-Warning "Oupps, $Collection collection does not exist"
    Break
}
 
# Create Software Update Deployment Package
if (Get-CMSoftwareUpdateDeploymentPackage -Name $ADRName)
{
    # All good
    Write-Output "$ADRName Software Update Deployment Package found, continue"
}
Else
{
    Write-Warning "$ADRName Software Update Deployment Package does not exist, create it"
    $NewDeploymentPackage = New-CMSoftwareUpdateDeploymentPackage -Name $ADRName -Path $DeployPackageLocation
}
 
# Distribute the Software Update Deployment Package
Start-CMContentDistribution -DeploymentPackageId $NewDeploymentPackage.PackageID -DistributionPointName $SiteServerFQDN
 
if (Get-CMSoftwareUpdateAutoDeploymentRule -Name $ADRName)
{
    # All good
    Write-Output "$ADRName Automatic Deployment Rule already exist"
 }
Else
{
    Write-Output "$ADRName Automatic Deployment Rule does not exist, create it"
     
    $Schedule = New-CMSchedule -DayOfWeek Monday -WeekOrder Third -Start ([Datetime]"08:00")
    New-CMSoftwareUpdateAutoDeploymentRule `
    -CollectionName $Collection `
    -DeploymentPackageName $ADRName `
    -Name $ADRName `
    -AddToExistingSoftwareUpdateGroup $False `
    -AlertTime 4 `
    -AlertTimeUnit Weeks `
    -AllowRestart $True `
    -AllowSoftwareInstallationOutsideMaintenanceWindow $True `
    -AllowUseMeteredNetwork $True `
    -AvailableTime 7 `
    -AvailableTimeUnit Days `
    -BulletinId "MS" `
    -DateReleasedOrRevised Last1month `
    -DeadlineImmediately $True `
    -DeployWithoutLicense $False `
    -DisableOperationManager $True `
    -DownloadFromInternet $True `
    -DownloadFromMicrosoftUpdate $True `
    -EnabledAfterCreate $True `
    -GenerateOperationManagerAlert $True `
    -GenerateSuccessAlert $True `
    -Language "English" `
    -LanguageSelection "English" `
    -NoInstallOnRemote $False `
    -NoInstallOnUnprotected $True `
    -Product $Products `
    -RunType RunTheRuleOnSchedule `
    -Schedule $Schedule `
    -SendWakeUpPacket $False `
    -Severity $Severity `
    -SuccessPercent 99 `
    -Superseded $False `
    -SuppressRestartServer $True `
    -SuppressRestartWorkstation $False `
    -UpdateClassification $UpdateClassifications `
    -UseBranchCache $False `
    -UserNotification DisplayAll `
    -UseUtc $True `
    -VerboseLevel AllMessages `
    -WriteFilterHandling $True `
}
 
    # Additional common parameters
    #-DeadlineTime $True `
    #-DeadlineTimeUnit Hours `
    #-MicrosoftAsVendor $True `
    #-ArticleId "100" `
    #-Location $DeployPackageLocation `
    #-AvailableImmediately $False `