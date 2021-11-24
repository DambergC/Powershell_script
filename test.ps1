# script to get dns-info for multiple ip-addresses or computernames

param ($List)

$ListToCheck = Get-Content $list


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
                    $props = @{'Data'=$data.Name
                               'Result'=$data.IPAddress}
                }

            else 
                {
                    $data = Resolve-DnsName $item                   
                    $props = @{ 'Data'=$item
                                'Result'=$data.NameHost}
                }
        $obj = New-Object -TypeName PSobject -Property $props
        Write-Output $obj
        }
}


