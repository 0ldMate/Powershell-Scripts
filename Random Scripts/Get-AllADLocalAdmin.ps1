$Computers = Get-ADComputer -Filter {OperatingSystem -like "Windows*"} -Properties cn | Where-Object { $_.enabled } | Sort-Object
foreach ($computer in $computers){
$a = if(Test-Connection -ComputerName $Computer.Name -Count 3 -Quiet){
$computer.name
$admins = Get-WmiObject win32_groupuser –computer $computer.name -ErrorVariable WMIError -ErrorAction SilentlyContinue
$admins = $admins |Where-Object {$_.groupcomponent –like '*"Administrators"'}  

if($WMIError){
    "$($computer.Name) errored with the RPC Server"
}
else{  
$admins |ForEach-Object {  
$_.partcomponent –match “.+Domain\=(.+)\,Name\=(.+)$” > $nul  
$matches[1].trim('"') + “\” + $matches[2].trim('"')  

}
}
"`n" 
Clear-Variable -Name admins
if($WMIError){Clear-Variable -Name WMIError}
}
else{
"$($computer.Name) is unreachable"
"`n" 
}
$a | Out-File -FilePath C:\Users\user\Desktop\SortedAdmins.txt -Append
}