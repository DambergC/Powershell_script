$OSInfo = Get-CimInstance -Class Win32_OperatingSystem
$languagePacks = $OSInfo.MUILanguages
$languagePacks