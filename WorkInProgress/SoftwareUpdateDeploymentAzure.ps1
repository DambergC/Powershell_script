Connect-AzAccount
Get-AzSubscription
Set-AzContext -Subscription '454435d2-36c7-4e0e-831f-673a818cc445'


Connect-AzAccount -TenantId bb63674f-acb1-488c-8cbe-4a83fb31f56a

$subscriptionId = "454435d2-36c7-4e0e-831f-673a818cc445"
$resourceGroupName = "Server-produktion"
$automationAccountName = "UpdateManager"
$authHeaders = @{
     'Content-Type'  = 'application/json'
     'Authorization' = 'Bearer ' + (Get-AzAccessToken).Token
 }
$result = Invoke-RestMethod -Uri https://management.azure.com/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.Automation/automationAccounts/$automationAccountName/softwareUpdateConfigurations?api-version=2017-05-15-preview -Headers $authHeaders
$result.value.id | ForEach-Object { Remove-AzResource -ResourceId $_ -Force }



$autoacc = Get-AzAutomationAccount -ResourceGroupName 'server-produktion' -Name 'UpdateManager'

$cfgs = $autoacc | Get-AzAutomationSoftwareUpdateConfiguration

Foreach ($cfg in $cfgs) { $cfg | Remove-AzAutomationSoftwareUpdateConfiguration }



$tenantid = (Get-AzContext).Tenant.Id
$subs = Get-AzSubscription | where {$_.TenantId -eq $tenantid}
$scope = @()
Foreach ($sub in $subs) { $scope += "/subscriptions/" + $sub.Id }
$query =  $autoacc | New-AzAutomationUpdateManagementAzureQuery -Scope $scope




$year = (Get-Date).Year
$month = (Get-Date).Month
$duration = New-TimeSpan -Hours 5
$time = "19:00"
$days = 3

while($month -le 12) {
  $day1 = [datetime]($month.ToString().PadLeft(2,'0') + "/01/" + $year.ToString() + " " + $time)
  $patchtues = (0..30 | % {$day1.adddays($_) } | ? {$_.dayofweek -like "Tue*"})[1]
  $winschname = $year.ToString() + "_" + $month.ToString().PadLeft(2,'0') + "_windows"
  $linschname = $year.ToString() + "_" + $month.ToString().PadLeft(2,'0') + "_linux"
  $schstart = $patchtues.AddDays($days)
  #Adjust for BST because Azure portal doesn't handle it
  if ((Get-Date -Date $schstart).IsDaylightSavingTime()) { $schstart = $schstart.AddHours(1) }
  $winsch = $autoacc | New-AzAutomationSchedule -Name $winschname -StartTime $schstart -TimeZone "GMT Standard Time" -OneTime -ForUpdateConfiguration
  $wincfg = $autoacc | New-AzAutomationSoftwareUpdateConfiguration -Windows -Schedule $winsch -AzureQuery $query -IncludedUpdateClassification Critical, Security -Duration $duration -RebootSetting IfRequired
  $linsch = $autoacc | New-AzAutomationSchedule -Name $linschname -StartTime $schstart -TimeZone "GMT Standard Time" -OneTime -ForUpdateConfiguration
  $lincfg = $autoacc | New-AzAutomationSoftwareUpdateConfiguration -Linux -Schedule $linsch -AzureQuery $query -IncludedPackageClassification Critical, Security -Duration $duration -RebootSetting IfRequired
  $month++
}