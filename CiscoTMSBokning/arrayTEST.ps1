 param(
       [Parameter(Position=0)][array]$Vlan
    )



 $vlan.ForEach(
 {
  $test = $_
  write-host -BackgroundColor Green -ForegroundColor white $test
 }
 )