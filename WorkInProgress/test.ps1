# definition of parameters
Param(
    [string]$StringParameter,
    [bool]$BooleanParameter = $False,
    [int]$IntegerParameter = 10,
    [switch]$SwitchParameter
)

# output of the parameter values
Write-host "StringParameter: $StringParameter"
Write-host "BooleanParameter: $BooleanParameter"
Write-host "IntegerParameter: $IntegerParameter"
Write-host "SwitchParameter: $SwitchParameter"