$Session0365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Office365UserCredential -Authentication  Basic -AllowRedirection
Import-PSSession $Session0365 -AllowClobber

        #gets the list of groups that are in the cloud 
$Groups = Get-DistributionGroup | select -ExpandProperty Name
foreach($g in $Groups) {
    $members = Get-DistributionGroupMember -Identity $g | Select-Object -ExpandProperty Name 
    $Report = Foreach ($m in $members){
    [PSCustomObject]@{
        GroupName = $g
        Members = $m
    }
}
    $Report | Export-Csv -Path $Env:USERPROFILE\Office365Groups.csv -Append -NoTypeInformation
}