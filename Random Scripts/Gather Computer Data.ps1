$ArrComputers =  "Computer"
$Creds = Get-Credential -UserName Admin -Message "Admin Pass"
#Specify the list of PC names in the line above. "." means local system

Clear-Host
foreach ($Computer in $ArrComputers) 
{
   if (Test-Connection -ComputerName $Computer -Count 3 -Quiet) {
   
    $computerSystem = get-wmiobject Win32_ComputerSystem -Computer $Computer -Credential $Creds
    $computerBIOS = get-wmiobject Win32_BIOS -Computer $Computer -Credential $Creds
    $computerOS = get-wmiobject Win32_OperatingSystem -Computer $Computer -Credential $Creds
    $computerCPU = get-wmiobject Win32_Processor -Computer $Computer -Credential $Creds
    $computerHDD = Get-WmiObject Win32_LogicalDisk -ComputerName $Computer -Filter drivetype=3 -Credential $Creds
    $mac = get-wmiobject -class "Win32_NetworkAdapterConfiguration" -computername $Computer -Credential $creds |Where{$_.IpEnabled -Match "True"} 
    $ip = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer -Credential $Creds | where { $_.ipaddress -like "1*" } | select -ExpandProperty ipaddress | select -First 1 
        write-host "System Information for: " $computerSystem.Name -BackgroundColor DarkCyan
        "-------------------------------------------------------"
        "Manufacturer: " + $computerSystem.Manufacturer
        "Model: " + $computerSystem.Model
        "SerialNumber: " + $computerBIOS.SerialNumber
        "CPU: " + $computerCPU.Name
        "HDD Capacity: "  + [int](($computerhdd | Measure-Object Size -Sum).Sum/1GB) + "GB"
        "RAM: " + "{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB) + "GB"
        "OS: " + $computerOS.caption + ", Service Pack: " + $computerOS.ServicePackMajorVersion
        "User logged In: " + $computerSystem.UserName
        "Last Reboot: " + $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
        "IPAddress: " + $ip
        "MACAddress: " + $mac.MACAddress
        ""
        "-------------------------------------------------------"
   }
}