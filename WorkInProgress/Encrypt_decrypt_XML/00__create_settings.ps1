
################################################
#
# SCRIPT ROOT
#
################################################

# Load scriptpath
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
    $scriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
} else {
    $scriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0])
}


Set-Location -Path $scriptPath

################################################
#
# FUNCTIONS
#
################################################

# load all functions
#. ".\epi__0__functions.ps1"


################################################
#
# CREATE ENCRYPTION KEYS
#
################################################

# create encryption keys
$cspParams = New-Object "System.Security.Cryptography.CspParameters"
$cspParams.KeyContainerName = "XML_ENC_RSA_KEY"
$rsaKey = [System.Security.Cryptography.RSACryptoServiceProvider]::new($cspParams)



################################################
#
# SETTINGS
#
################################################


#$pass = Read-Host -AsSecureString "Please enter the password for epi"
#$passEncrypted = Get-PlaintextToSecure ((New-Object PSCredential "dummy",$pass).GetNetworkCredential().Password)

$settings = @{
    keyContainerName = $cspParams.KeyContainerName
    keyName = "rsaKey"
    elementsToEncrypt = @("Password", "ConnectionString","PeopleStageConnectionString")
    keySize = 256
    #logfile = "D:\Apteco\Log\ps_custom__epi_transform.log"
}


################################################
#
# PACK TOGETHER SETTINGS AND SAVE AS JSON
#
################################################

# create json object
$json = $settings | ConvertTo-Json -Depth 8 # -compress

# print settings to console
$json

# save settings to file
$json | Set-Content -path "$( $scriptPath )\settings.json" -Encoding UTF8

