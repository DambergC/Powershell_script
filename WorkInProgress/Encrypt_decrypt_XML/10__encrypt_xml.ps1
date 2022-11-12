<#

Reference: 
https://docs.microsoft.com/de-de/dotnet/standard/security/how-to-encrypt-xml-elements-with-symmetric-keys
https://social.msdn.microsoft.com/Forums/de-DE/a7732f36-ae89-471d-b5de-af84b2d14cca/problem-beim-entschlsseln-von-xml-elementen


#>



################################################
#
# PATH
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
# SETTINGS
#
################################################


# Load settings
$settings = Get-Content -Path "$( $scriptPath )\settings.json" -Encoding UTF8 -Raw | ConvertFrom-Json


# settings
$inputfile = ".\epi__responses_configuration.xml"
$outputfile = ".\epi__responses_configuration_2.xml"

$elementsToEncrypt = $settings.elementsToEncrypt



################################################
#
# FUNCTIONS / ASSEMBLIES
#
################################################

# load all functions
#. ".\epi__0__functions.ps1"

# load assemblies
Add-Type -AssemblyName System.Security #, System.Text.Encoding


################################################
#
# ENCRYPTION KEYS
#
################################################

# create encryption keys
$cspParams = New-Object "System.Security.Cryptography.CspParameters"
$cspParams.KeyContainerName = $settings.keyContainerName
$rsaKey = [System.Security.Cryptography.RSACryptoServiceProvider]::new($cspParams)
$keyName = $settings.keyName

# create encryption key
$sessionKey = new-object "System.Security.Cryptography.RijndaelManaged"
$sessionKey.KeySize = $settings.keySize


################################################
#
# ENCRYPTION XML PARTS
#
################################################

# load xml file
$xml = New-Object "xml"
$xml.PreserveWhitespace
$xml.Load($inputfile)

$elementsToEncrypt | ForEach {

    $tag = $_

    # select the xml element
    $xmlElement = $xml.GetElementsByTagName($tag)[0]

    # encrypt xml element
    $eXml = new-object "System.Security.Cryptography.Xml.EncryptedXml"
    $encryptedElement = $eXml.EncryptData($xmlElement,$sessionKey,$false)

    # describe encryption
    $edElement = New-Object "System.Security.Cryptography.Xml.EncryptedData"
    $edElement.Type = [System.Security.Cryptography.Xml.EncryptedXml]::XmlEncElementUrl
    $edElement.Id = $tag
    $edElement.EncryptionMethod = [System.Security.Cryptography.Xml.EncryptedXml]::XmlEncAES256Url
    
    # create encrypted key
    $ek = New-Object "System.Security.Cryptography.Xml.EncryptedKey"
    $encryptedKey = [System.Security.Cryptography.Xml.EncryptedXml]::EncryptKey($sessionKey.Key, $rsaKey, $false)
    $ek.CipherData = [System.Security.Cryptography.Xml.CipherData]::new($encryptedKey)
    $ek.EncryptionMethod = [System.Security.Cryptography.Xml.EncryptionMethod]::new([System.Security.Cryptography.Xml.EncryptedXml]::XmlEncRSA15Url)
    
    # create data reference
    $dRef = New-Object "System.Security.Cryptography.Xml.DataReference"
    $dRef.Uri = "#" + $tag
    $ek.AddReference($dRef)
    $edElement.KeyInfo.AddClause([System.Security.Cryptography.Xml.KeyInfoEncryptedKey]::new($ek))
    
    # create key info
    $kin = New-Object "System.Security.Cryptography.Xml.KeyInfoName"
    $kin.Value = $keyName
    $ek.KeyInfo.AddClause($kin)
    $edElement.CipherData.CipherValue = $encryptedElement
    
    # replace element
    [System.Security.Cryptography.Xml.EncryptedXml]::ReplaceElement($xmlElement,$edElement,$false)

}

# save the xml
$xml.Save($outputfile)


# finally
$sessionKey.Clear()
