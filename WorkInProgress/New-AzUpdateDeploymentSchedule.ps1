Function New-AzUpdateDeploymentSchedule {
    <#
.SYNOPSIS
Schedule Azure updates
.DESCRIPTION
This cmdlet combines the three cmdlets that are needed to create a update schedule.
New-AzAutomationSchedule
New-AzAutomationUpdateManagementAzureQuery
New-AzAutomationSoftwareUpdateConfiguration
Created for Azure Update Automation.
.PARAMETER StartTime
DateTimeOffset object to define the time updates should start
.PARAMETER ResourceGroupName
The resourcegroup where the automation account is placed
.PARAMETER AutomationAccountName
The name of the automation account
.PARAMETER LogAnalyticsWorkspace
The name of the log analytics workspace
.PARAMETER UpdateClassification
What classifications need to be updated.
Defaults to selects all classifications.
.PARAMETER AzureQuery
Switch to show if you need an AzureQuery
.PARAMETER Scope
Only if you need a Azure Query: The scope for the updates.
Defaults to the subscription of the subscription
.PARAMETER Location
Only if you need an Azure Query:
Only update the Azure VMs with this location
.PARAMETER Tags
Only if you need an Azure Query:
Only update Azure VMs with these tags
.PARAMETER NonAzureQueryName
When using Non-Azure VMs, pass the query name
.EXAMPLE
$Parameters = @{
    StartTime              = (Get-Date -Date "13-07-2020 08:00")
    ResourceGroupName      = "ResourceGroupName"
    AutomationAccountName  = "AutomationAccountName"
    LogAnalyticsWorkspace = "4besloganalyticsworkspace"
    AzureQuery             = $true
    Location               = "westeurope"
    Tags                   = @{"example" = "true" }
}
New-AzUpdateDeploymentSchedule @Parameters
.EXAMPLE
$Parameters = @{
    StartTime              = (Get-Date -Date "02-08-2020 08:00")
    ResourceGroupName      = "ResourceGroupName"
    AutomationAccountName  = "AutomationAccountName"
    LogAnalyticsWorkspace  = "4besloganalyticsworkspace"
    NonAzureSavedQueryName = "Example"
}
New-AzUpdateDeploymentSchedule @Parameters
.NOTES
Created by Barbara Forbes
@ba4bes
https://4bes.nl
#>
    [CmdletBinding(DefaultParameterSetName = "__AllParameterSets")]
    param(
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "__AllParameterSets"
        )]
        [System.DateTimeOffset]$StartTime ,
        [Parameter(ParameterSetName = "__AllParameterSets")]
        [string]$ResourceGroupName ,
        [Parameter(ParameterSetName = "__AllParameterSets")]
        [string]$AutomationAccountName ,
        [Parameter(ParameterSetName = "__AllParameterSets")]
        [string]$LogAnalyticsWorkspace ,
        [Parameter(ParameterSetName = "__AllParameterSets")]
        [string[]]$UpdateClassification = @(
            "Unclassified"
            "Critical"
            "Security"
            "UpdateRollup"
            "FeaturePack"
            "ServicePack"
            "Definition"
            "Tools"
            "Updates"
        ),
        [Parameter(ParameterSetName = "__AllParameterSets")]
        [string]$NonAzureSavedQueryName,
        [Parameter(ParameterSetName = "__AzureSavedQuery")]
        [switch]$AzureQuery ,
        [Parameter(ParameterSetName = "__AzureSavedQuery")]
        [array]$Scope = @("/subscriptions/$((Get-AzContext).Subscription.Id)"),
        [Parameter(ParameterSetName = "__AzureSavedQuery")]
        [string]$Location ,
        [Parameter(ParameterSetName = "__AzureSavedQuery")]
        [hashtable]$Tags
    )

    # Create the time based on the schedule date and schedule time
    $ScheduleTime = Get-Date $StartTime.LocalDateTime -Format "yyyy-MM-dd"
    # Create the parameters for setting the schedule
    $ScheduleParameters = @{
        ResourceGroupName     = $ResourceGroupName
        AutomationAccountName = $AutomationAccountName
        name                  = $ScheduleTime
        StartTime             = $StartTime
        OneTime               = $True
    }
    # Create the schedule that will be used for the updates
    Try {
        $AutomationSchedule = New-AzAutomationSchedule @ScheduleParameters -ErrorAction Stop -Verbose
        Write-Verbose "Schedule has been created"
    }
    Catch {
        Throw "Could not create Automation Schedule: $_"
    }

    # create the parameters for the updateschedule
    $UpdateParameters = @{
        ResourceGroupName            = $ResourceGroupName
        AutomationAccountName        = $AutomationAccountName
        Schedule                     = $AutomationSchedule
        Windows                      = $true
        Duration                     = New-TimeSpan -Hours 2
        RebootSetting                = "Always"
        IncludedUpdateClassification = $UpdateClassification
    }
    if ($NonAzureSavedQueryName) {
        # Create object for query
        $nonAzQuery = [Microsoft.Azure.Commands.Automation.Model.UpdateManagement.NonAzureQueryProperties]::New()
        $nonAzQuery.FunctionAlias = $NonAzureSavedQueryName
        $nonAzQuery.WorkspaceResourceId = (Get-AzResource -Name $LogAnalyticsWorkspace).ResourceId

        $UpdateParameters.Add("NonAzureQuery", $nonAzQuery)
    }
    if ($AzureQuery) {
        $QueryParameters = @{
            ResourceGroupName     = (Get-AzResource -Name $LogAnalyticsWorkspace).ResourceGroupName
            AutomationAccountName = $AutomationAccountName
            Scope                 = $Scope
        }
        if ($Location) {
            $QueryParameters.Add("Location", $Location)
        }
        if ($Tags) {
            $QueryParameters.Add("Tag", $Tags)
        }
        Try {
            $AzQuery = New-AzAutomationUpdateManagementAzureQuery @QueryParameters -Verbose -ErrorAction Stop
            Write-Verbose "Query has been created"
        }
        Catch {
            Throw "Could not create query: $_"
        }
        $UpdateParameters.Add("AzureQuery", $AzQuery)
    }
    # Set the Automation Schedule
    Try {
        New-AzAutomationSoftwareUpdateConfiguration @UpdateParameters -Verbose -ErrorAction Stop
    }
    Catch {
        Throw "Could not create update schedule: $_"
    }
}