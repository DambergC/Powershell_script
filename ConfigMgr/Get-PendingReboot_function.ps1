Function Get-PendingReboot
{ 

<#
 
.SYNOPSIS
    Gets the pending reboot status on a local computer. Return
 
.DESCRIPTION
    Queries the registry and WMI to determine if the system waiting for a reboot, from:
        CBServicing = Component Based Servicing (Windows 2008)
        WindowsUpdate = Windows Update / Auto Update (Windows 2003 / 2008)
        CCMClientSDK = SCCM 2012 Clients only (DetermineIfRebootPending method) otherwise $null value
        PendFileRename = PendingFileRenameOperations (Windows 2003 / 2008)
 
    Returns hash table similar to this:
 
    Computer : MYCOMPUTERNAME
    LastBootUpTime : 01/12/2014 11:53:04 AM
    CBServicing : False
    WindowsUpdate : False
    CCMClientSDK : False
    PendFileRename : False
    PendFileRenVal :
    RebootPending : False
    ErrorMsg :
 
    NOTES:
    ErrorMsg only contains something if an error occured
 
.EXAMPLE
    Get-PendingReboot
 
.EXAMPLE
    $PRB=Get-PendingReboot
    $PRB.RebootPending
 
.NOTES
    Based On: http://gallery.technet.microsoft.com/scriptcenter/Get-PendingReboot-Query-bdb79542
#>
 
[CmdletBinding()] 

param
( 

  [Parameter(Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)] 
  [String[]]
  $ComputerName="$env:COMPUTERNAME", 

  [String]
  $ErrorLog,

  [Parameter()]
  [System.Management.Automation.PSCredential]
  [System.Management.Automation.Credential()]
  $Credential = [System.Management.Automation.PSCredential]::Empty,
  
  [Switch]
  $Debugging

  ) 
 
Begin {  }## End Begin Script Block
Process { 
  Foreach ($Computer in $ComputerName) { 
  
Try { 
      ## Setting pending values to false to cut down on the number of else statements
      $CompPendRen,$PendFileRename,$Pending,$SCCM = $false,$false,$false,$false 
       
      ## Setting CBSRebootPend to null since not all versions of Windows has this value
      $CBSRebootPend = $null 
             
      ## Querying WMI for build version
      $WMI_OS = Get-WmiObject -Class Win32_OperatingSystem -Property BuildNumber, CSName -ComputerName $Computer -Credential $Credential -ErrorAction Stop 
 
      ## Making registry connection to the local/remote computer
      $HKLM = [UInt32] "0x80000002" 
      
      #$WMI_Reg = [WMIClass] "\\$Computer\root\default:StdRegProv" #Commenting for now while I rework the credentials/WMI call.


<#### Adding test for WMI reg calls#>

      $WMI_Reg = Get-Wmiobject -list "StdRegProv" -namespace root\default -Computername $computer -Credential $Credential
      #$value = $wmi_reg.GetStringValue($HKEY_Local_Machine,$key,$valuename).svalue


      ## If Vista/2008 & Above query the CBS Reg Key
      If ([Int32]$WMI_OS.BuildNumber -ge 6001) { 
        $RegSubKeysCBS = $WMI_Reg.EnumKey($HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\") 
        $CBSRebootPend = $RegSubKeysCBS.sNames -contains "RebootPending"     
      } 
               
      ## Query WUAU from the registry
      $RegWUAURebootReq = $WMI_Reg.EnumKey($HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\") 
      $WUAURebootReq = $RegWUAURebootReq.sNames -contains "RebootRequired" 
             
      ## Query PendingFileRenameOperations from the registry
      $RegSubKeySM = $WMI_Reg.GetMultiStringValue($HKLM,"SYSTEM\CurrentControlSet\Control\Session Manager\","PendingFileRenameOperations") 
      $RegValuePFRO = $RegSubKeySM.sValue 
 
      ## Query ComputerName and ActiveComputerName from the registry
      $ActCompNm = $WMI_Reg.GetStringValue($HKLM,"SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName\","ComputerName")       
      $CompNm = $WMI_Reg.GetStringValue($HKLM,"SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\","ComputerName") 
      If ($ActCompNm -ne $CompNm) { 
    $CompPendRen = $true 
      } 
             
      ## If PendingFileRenameOperations has a value set $RegValuePFRO variable to $true
      If ($RegValuePFRO) { 
        $PendFileRename = $true 
      } 
 
      ## Determine SCCM 2012 Client Reboot Pending Status
      ## To avoid nested 'if' statements and unneeded WMI calls to determine if the CCM_ClientUtilities class exist, setting EA = 0
      $CCMClientSDK = $null 
      $CCMSplat = @{ 
        NameSpace='ROOT\ccm\ClientSDK' 
        Class='CCM_ClientUtilities' 
        Name='DetermineIfRebootPending' 
        ComputerName=$Computer
        Credential=$Credential 
        ErrorAction='Stop' 
        } 
      ## Try CCMClientSDK
      Try { 
    $CCMClientSDK = Invoke-WmiMethod @CCMSplat 
      } Catch [System.UnauthorizedAccessException] { 
    $CcmStatus = Get-Service -Name CcmExec -ComputerName $Computer -ErrorAction SilentlyContinue 
    If ($CcmStatus.Status -ne 'Running') { 
        #Write-Warning "$Computer`: Error - CcmExec service is not running."
        $CCMClientSDK = $null 
    } 
      } Catch { 
    $CCMClientSDK = $null 
      } 
 
      If ($CCMClientSDK) { 
    If ($CCMClientSDK.ReturnValue -ne 0) { 
      Write-Warning "Error: DetermineIfRebootPending returned error code $($CCMClientSDK.ReturnValue)"     
        } 
        If ($CCMClientSDK.IsHardRebootPending -or $CCMClientSDK.RebootPending) { 
      $SCCM = $true 
        } 
      } 
       
      Else { 
    $SCCM = $null 
      } 
 
      ## Creating Custom PSObject and Select-Object Splat
      $SelectSplat = @{ 
    Property=( 
        'Computer', 
        'CBServicing', 
        'WindowsUpdate', 
        'CCMClientSDK', 
        'PendComputerRename', 
        'PendFileRename', 
        'PendFileRenVal', 
        'RebootPending' 
    )} 

      $results = New-Object -TypeName PSObject -Property @{ 
    Computer=$WMI_OS.CSName 
    CBServicing=$CBSRebootPend 
    WindowsUpdate=$WUAURebootReq 
    CCMClientSDK=$SCCM 
    PendComputerRename=$CompPendRen 
    PendFileRename=$PendFileRename 
    PendFileRenVal=$RegValuePFRO 
    RebootPending=($CompPendRen -or $CBSRebootPend -or $WUAURebootReq -or $SCCM -or $PendFileRename) 
      } | Select-Object @SelectSplat 


if ($debugging) {$results}

elseIf ($results.rebootpending -eq $true) {
    Return "Pending Reboot"
    }
elseif ($results.rebootpending -eq $false) {
    #Return "$results"
    Return "False"
    }
else {
    Return "Unknown"
    }
}

  
Catch { 
      #Write-Warning "$Computer`: $_"
      Return $_
      ## If $ErrorLog, log the file to a user specified location/path
      If ($ErrorLog) { 
    Out-File -InputObject "$Computer`,$_" -FilePath $ErrorLog -Append 
      }         
  }       

  }## End Foreach ($Computer in $ComputerName)
}## End Process
 
End {  }## End End
 
}## End Function Get-PendingReboot