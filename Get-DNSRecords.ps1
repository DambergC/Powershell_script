#########################################################
#.Synopsis
#  Read txt-file and run multiple check on DNS
#.DESCRIPTION
#  The script import a txt-file with multiple values to check against DNS 
#.EXAMPLE
#   Get-DNSRecords.ps1 -List .\textfile.txt
#
#   Txt-file strukture:(No header!)
#
#   dc01
#   192.168.3.214
#
#   Result on screen:
#
#   Result                   Data
#   ------                   ----
#   192.168.3.200            dc01.corp.viamonstra.com
#   CM01.corp.viamonstra.com 192.168.3.214
#
#   Skript created by Christian Damberg
#   christian@damberg.org
#   https://www.damberg.org
#
#   Scriptversion 1.0
#   2021-11-09
#
#########################################################

param ($List)

$ListToCheck = Get-Content $list
$Result = @()

foreach ($item in $ListToCheck) 
{
    $communication = Test-Connection $item -Count 1 -Quiet

  if ($communication -eq $true)  
        {
        $GetDnsType = Resolve-DnsName $item
        $DnsType = $GetDnsType.type
       
            if ($DnsType -eq 'A') 
                {
                    $data = Resolve-DnsName $item  -Type A
                    $props = @{ 'Result'=$data.IPAddress
                                'Data'=$data.Name
                               }
                }

            else 
                {
                    $data = Resolve-DnsName $item                   
                    $props = @{ 'Result'=$data.NameHost
                                'Data'=$item
                                }
                }
        $obj = New-Object -TypeName PSobject -Property $props
        $Result += $obj
        }

       else 
        {
          $props = @{ 'Result'='No response'
                      'Data'=$item
        }
          $obj = New-Object -TypeName PSobject -Property $props
          $Result += $obj
       } 
}

$Result | Format-Table