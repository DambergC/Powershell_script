#Specify ISO Source location

#$FoD_Source = "$env:USERPROFILE\Downloads\W10RSAT_FOD\1903\1903_FoD_Disk1.iso"
$FoD_Source = "F:\ISO\FEATURE ON DEMAND\en_windows_10_features_on_demand_part_1_version_1903_x64_dvd_1076e85a.iso"
#Mount ISO

Mount-DiskImage -ImagePath "$FoD_Source"

$path = (Get-DiskImage "$FoD_Source" | Get-Volume).DriveLetter

#Language desired
$lang2 ='en-us'
$lang = "*en-us*"

#folder 

$dest = New-Item -ItemType Directory -Path "$env:SystemDrive\temp\TextToSpeech_$lang2" -force

#get RSAT files 

Get-ChildItem ($path+":\") -name -recurse -include *~amd64~~.cab,*~wow64~~.cab,*~amd64~$lang~.cab,*~wow64~$lang~.cab -exclude *InternationalFeatures*,*webdriver*,*Accessibility*,*onecore*,*internetexp*,*ipam*,*irda*,*snmp*,*print*,*tablet*,*ras*,*server*,*remote*,*Holographic*,*NetFx3*,*OpenSSH*,*Msix*,*xps*,*storage*,*tools* -filter $lang |
ForEach-Object {copy-item -Path ($path+“:\”+$_) -Destination $dest.FullName -Force -Container}

#get metadata

copy-item ($path+":\metadata") -Destination $dest.FullName -Recurse -Filter $lang

copy-item ($path +“:\"+“FoDMetadata_Client.cab”) -Destination $dest.FullName -Force -Container

#Dismount ISO

Dismount-DiskImage -ImagePath "$FOD_Source"