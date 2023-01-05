
# Get computer distinguishedname
$res=Get-ADComputer -Identity cleint01 | Select-Object DistinguishedName -ExpandProperty Distinguishedname

# Remove computername from distinguishedname
$clean = $res.DistinguishedName.Replace("CN=CLEINT01,","")

# write OU path to adminDescription
Set-ADComputer -Identity cleint01 -Add @{adminDescription="$clean"} -Verbose

# Get path from adminDescription
$adpath = Get-ADComputer -Identity cleint01 -Properties * | select-object adminDescription -ExpandProperty adminDescription

# Move object to disabled computers
Move-ADObject -Identity $res -TargetPath 'OU=Disabled Computers,OU=ViaMonstra,DC=corp,DC=viamonstra,DC=com'

# get new path
$res2=Get-ADComputer -Identity cleint01 | Select-Object DistinguishedName -ExpandProperty Distinguishedname

# restore to old path
Move-ADObject -Identity $res2 -TargetPath $adpath

