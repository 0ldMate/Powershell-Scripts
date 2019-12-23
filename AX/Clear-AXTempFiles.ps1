Import-Module Active Directory
$Computers = Get-ADComputer -Filter {OperatingSystem -like "Windows*"} -Properties cn | Where-Object { $_.enabled }
foreach ($computer in $computers){
if(Test-Connection -ComputerName $Computer.Name -Count 3 -Quiet){
    $Users = Get-ChildItem -Path "\\$Computer\C$\Users"
    foreach($user in $users){
        Get-ChildItem -Path "\\$Computer\C$\Users\$User\Appdata\local\*" -Include *.auc | Remove-Item
}
}
else{
"$computer.Name is unreachable"
}
    
}