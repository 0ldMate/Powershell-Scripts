Import-Module Active Directory
$Computers = Get-ADComputer -Filter {OperatingSystem -like "Windows Server*"} -Properties cn | Where-Object { $_.enabled }
foreach ($computer in $computers){
$Report = if(Test-Connection -ComputerName $Computer.Name -Count 3 -Quiet){
Write-Host "$Computer Services"
Get-WmiObject win32_service -computer $Computer.Name -ErrorAction SilentlyContinue | Where-Object {$_.startname -like "COMPANY*"} | Select-Object __SERVER,Name, StartName
}
else{
"$computer.Name is unreachable"
}
$Report | Out-File $env:USERPROFILE\Desktop\ServiceUsers.csv -Append
}