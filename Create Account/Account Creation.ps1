$UserCredential = Get-Credential -UserName Administrator -Message "Administrator Account"
$JiraCredential = Get-Credential -UserName JiraAdmin -Message "Jira account eg JiraAdmin"
. "C:\Program Files\Microsoft Dynamics AX\60\ManagementUtilities\Microsoft.Dynamics.ManagementUtilities.ps1"
$SessionDC = New-PSSession -ComputerName "DC" -Credential $UserCredential -Name DC
Import-Module -PSSession $SessionDC -Name ActiveDirectory
Import-Module JiraPS
Set-JiraConfigServer -Server "http://help.jira.com.au"
New-JiraSession -Credential $JiraCredential
Import-Csv "$PSScriptRoot\Employee.csv" | ForEach-Object {
$Name = $_.Name
$Firstname = $Name.Split(" ")[0]
$Surname = $Name.Split(" ")[1]
$EmailAddress = "$Firstname" + ".$Surname" + "@Company.com"
$Title = $_.Title
$CopyUser = $_.CopyUser
$Description = $_.Title
$MobileNumber = $_.Mobilephone
$PhoneNumber = $_.OfficePhone
$DisplayName = $Name
$SAMAccountName = $Firstname.Substring(0,1) + $Surname
$Company = "Company"
$Fax = $_.Fax
$StreetAddress = "123 Fake Street"
$City = "Somewhere"
$PostalCode = "0000"
$State = "Somewhere"
$Country = "Country"
$ScriptPath = "Test.bat"
$HomePage = "Company.com.au"
$UserPrincipal = "$SAMAccountName" + "@Company.com"
$TicketNumber = $_.TicketNumber

$TargetUser = Get-ADUser -Identity $CopyUser -Properties *
$TargetOU = $TargetUser.DistinguishedName -replace "CN=$($TargetUser.CN),"
$TargetGroups = $TargetUser.MemberOf
$TargetManager = $TargetUser.Manager
$TargetDepartment = $TargetUser.Department

$pass = Read-Host -Prompt "Input password"
New-ADUser -Name $Name -GivenName $FirstName -Surname $Surname -EmailAddress $EmailAddress -Title $Title -Department $TargetDepartment -MobilePhone $MobileNumber -OfficePhone $PhoneNumber -DisplayName $DisplayName -SamAccountName $SAMAccountName -Manager $TargetManager -Company $Company -Fax $Fax -StreetAddress $StreetAddress -City $City -PostalCode $PostalCode -State $State -Country $Country -ScriptPath $ScriptPath -HomePage $HomePage -UserPrincipalName $UserPrincipal -Description $Description
$TargetGroups | Add-ADGroupMember -Members $SAMAccountName
$NewObjectID = Get-ADUser -Identity $SAMAccountName | Select-Object -ExpandProperty ObjectGUID
Move-ADObject -Identity $NewObjectID -TargetPath $TargetOU
Set-ADAccountPassword $SAMAccountName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$pass" -force)
Set-ADUser -Identity $SAMAccountName -Enabled $true
Write-Host "$Name's AD account has been created"

$Ticket = "JIRA-" + $TicketNumber

$JiraComment = @"
$Name's account has been created. 
Once the account has been set up on the machine, run the Password reset script.

*If something is incorrect, this is all automatic from account creation script.*
"@

$JiraWorkLog = @"
Account has been created by script.
Username is $SAMAccountName
Password is $pass
"@

Add-JiraIssueComment -Issue "$Ticket" -Comment "$JiraComment"  
Add-JiraIssueWorklog -Issue $Ticket -TimeSpent "00:15" -Comment $JiraWorkLog -DateStarted (Get-Date)

$AXUsername = if ($SAMAccountName.length -gt 5){
    $SAMAccountName.SubString(0,5)
}
else {
    $SAMAccountName
}

New-AXUser -AccountType WindowsUser -AXUserId $AXUsername -UserName $SAMAccountName -UserDomain corp.Company.com.au
$Roles = Get-AXSecurityRole -AxUserID $CopyUser -ErrorVariable $AXError | Select-Object -ExpandProperty AOTName
if (!$AXError) {
    $FixedCopy = Read-Host -Prompt "AX Username - $copyuser is invalid, please input the correct User ID"
    $Roles = Get-AXSecurityRole -AxUserID $FixedCopy | Select-Object -ExpandProperty AOTName
}
ForEach($Role in $Roles) {Add-AXSecurityRoleMember -AxUserID $AXUserName -AOTName $Role;Write-Host "Assigning $Role to $Name"}

}

Read-Host -Prompt "Account will be synced to O365 shortly"
