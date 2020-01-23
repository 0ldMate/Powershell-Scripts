$ArrComputers =  (Get-Content C:\comps.txt)
#not written by me, modified.

#Specify the list of PC names in the line above. "." means local system

Clear-Host
foreach ($Computer in $ArrComputers) 
{
   if (Test-Connection -ComputerName $Computer -Count 2 -Quiet) {
   
   

    $computerSystem = get-wmiobject Win32_ComputerSystem -Computer $Computer
    $computerBIOS = get-wmiobject Win32_BIOS -Computer $Computer
    $computerOS = get-wmiobject Win32_OperatingSystem -Computer $Computer
    $computerCPU = get-wmiobject Win32_Processor -Computer $Computer
    $computerHDD = Get-WmiObject Win32_LogicalDisk -ComputerName $Computer -Filter drivetype=3
    $mac = get-wmiobject -class "Win32_NetworkAdapterConfiguration" -computername $Computer |Where{$_.IpEnabled -Match "True"}    
    $ip = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer  | where { $_.ipaddress -like "1*" } | select -ExpandProperty ipaddress | select -First 1
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
 


#Build the CSV file
$csvObject = New-Object PSObject -property @{
    'PCName' = $computerSystem.Name
    'Manufacturer' = $computerSystem.Manufacturer
    'Model' = $computerSystem.Model
    'SerialNumber' = $computerBIOS.SerialNumber
    'RAM' = "{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB) + "GB"
    'HDDSize' = [int](($computerhdd | Measure-Object Size -Sum).Sum/1GB)
    'CPU' = $computerCPU.Name
    'OS' = $computerOS.caption
    'User' = $computerSystem.UserName
    'BootTime' = $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
    'IP' = $ip
    'MAC' = $mac.MACAddress 
    }

 
         
$csvObject | Select PCName, Manufacturer, Model, SerialNumber, User, CPU, RAM, OS, HDDSize, HDDFree, BootTime, IP, MAC | Export-Csv 'C:\export.csv' -NoTypeInformation -Append

}
    else{New-Object PSObject -property @{Computer=$Computer; Account="NOT_AVAILABLE"} | Export-Csv 'C:\offline.csv' -NoTypeInformation -Append



}
    

}