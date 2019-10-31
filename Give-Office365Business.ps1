
    # Script is written to change users from Business Essentials and ProPlus license to Business Premium

        #Connects to Microsoft Online
Connect-MsolService -Credential (Get-Credential)

            #These are used to gather info on who has what license
        #Get-MsolUser -all| select -ExpandProperty Licenses | where {($_.Licenses).AccountSkuID -match "Plus"}
        #Get-MsolUser | Where-Object {($_.licenses).AccountSkuid -match "Flow"}
        #Get-MsolUser -UserPrincipalName User@company.com.au

    #Script gathers info on people who have the ProPlus added to them, Removes both the essential and proplus license and adds the premium license

$EssentialUsers = Get-MsolUser -all| Where-Object {($_.Licenses).AccountSkuID -match "Company:OFFICESUBSCRIPTION"}
Foreach($EssentialUser in $EssentialUsers) {
    $Confirm = Read-Host -Prompt "Change $($EssentialUser.DisplayName)'s Business Essentials license to Business Premium (Yes/No)"
    if($Confirm -eq "Yes"){
    Set-MsolUserLicense -UserPrincipalName $EssentialUser.UserPrincipalName -RemoveLicenses "company:O365_BUSINESS_ESSENTIALS","company:OFFICESUBSCRIPTION" -AddLicenses "company:O365_BUSINESS_PREMIUM"
    }
}

#needs two modules install-module azuread and msonline