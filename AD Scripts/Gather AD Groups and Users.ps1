$ADGroupList = Get-ADGroup -Filter * | Where-Object{$_.Name -like "Marketing*"} | Select Name -ExpandProperty Name | Sort 
ForEach($Group in $ADGroupList) 
{ 
Write-Host "Group: "$Group
Get-ADGroupMember -Identity $Group | Select Name -ExpandProperty Name | Sort 
Write-Host "" 

}
