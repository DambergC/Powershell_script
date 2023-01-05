[xml]$file1 = Get-Content C:\MjölbyTEST\PROD\PPP_WEB_LON\web.config
[xml]$file2 = Get-Content C:\MjölbyTEST\Backup\PPP_WEB_LON\web.config
$results = Compare-Object ($file1.SelectNodes("//add[@key]") | Select-Object -ExpandProperty Key) ($file2.SelectNodes("//add[@key]")  | Select-Object -ExpandProperty Key) | Where-Object{ $_.SideIndicator -eq "=>" } | ForEach-Object{ $_.InputObject}


foreach ($result in $results)
{
  
  $file1.configuration.appSettings.add | Where-Object key -eq $result

  
 }

 

foreach ($result in $results)
{

  $file2.configuration.appSettings.add | Where-Object key -eq $result
  
 }






