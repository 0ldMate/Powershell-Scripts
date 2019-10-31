
        ###!!! SCIRPT INACTIVE. Script has been implemented into Account creation script
        

. "C:\Program Files\Microsoft Dynamics AX\60\ManagementUtilities\Microsoft.Dynamics.ManagementUtilities.ps1"



New-AXUser -AccountType WindowsUser -AXUserId Username -UserName Username -UserDomain corp.company.com.au -Company DAT
$Roles = Get-AXSecurityRole -AxUserID $CopyUser | Select-Object -ExpandProperty AOTName
ForEach($Role in $Roles) {Add-AXSecurityRoleMember -AxUserID username -AOTName $Role}