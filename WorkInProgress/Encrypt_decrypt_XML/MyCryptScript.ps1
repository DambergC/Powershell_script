
$elementsToEncryptSettings = 'InstansPassword'

$keySizeSettings = '256'
$keyNameSettings = "rsaKey"
$keyContainerNameSettings = "XML_ENC_RSA_KEY"

# Load scriptpath
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
    $scriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
} else {
    $scriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0])
}


Set-Location -Path $scriptPath


# settings
$inputfile = ".\BooksRecord.xml"
$outputfile = ".\BooksRecord2.xml"

$elementsToEncrypt = $elementsToEncryptSettings

# load assemblies
Add-Type -AssemblyName System.Security #, System.Text.Encoding

################################################
#
# ENCRYPTION KEYS
#
################################################

# create encryption keys
$cspParams = New-Object "System.Security.Cryptography.CspParameters"
$cspParams.KeyContainerName = $keyContainerNameSettings
$rsaKey = [System.Security.Cryptography.RSACryptoServiceProvider]::new($cspParams)
$keyName = $keyNameSettings

# create encryption key
$sessionKey = new-object "System.Security.Cryptography.RijndaelManaged"
$sessionKey.KeySize = $keySizeSettings

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
