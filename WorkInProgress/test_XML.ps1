
#skapa xml-dokument
$xmlWriter = New-Object System.XMl.XmlTextWriter("C:\Temp\BooksRecord.xml",$null)
$xmlWriter.Formatting = 'Indented'
$xmlWriter.Indentation = 1
$XmlWriter.IndentChar = "`t"

$xmlWriter.WriteStartDocument()

$xmlWriter.WriteStartElement("Configuration")  # Configuration Startnode

$xmlWriter.WriteElementString("CustomerBigram","BIGRAM")
$xmlWriter.WriteElementString("InstansUser","ViwInstall213123")
$xmlWriter.WriteElementString("InstansPassword","Visma2016!")
$xmlWriter.WriteEndElement() # Configuration endnode
$xmlWriter.Flush()
$xmlWriter.Close()


[XML]$Code = Get-Content C:\temp\BooksRecord.xml
$CodeSecureString = ConvertTo-SecureString $Code -AsPlainText -Force
$Encrypted = ConvertFrom-SecureString -SecureString $CodeSecureString
$Encrypted | Export-Clixml -Path C:\temp\BooksRecord.xml


[xml]$decrypt = Get-Content C:\temp\BooksRecord.xml

$MySecureString = ConvertTo-SecureString -String $decrypt -AsPlainText -Force


$Marshal = [System.Runtime.InteropServices.Marshal]
$Bstr = $Marshal::SecureStringToBSTR($MySecureString)
$test = $Marshal::PtrToStringAuto($Bstr)
Write-host $test
$Marshal::ZeroFreeBSTR($Bstr)