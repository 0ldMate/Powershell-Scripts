$UserCredential = Get-Credential -UserName Administrator -Message "Input Admin Account Creds"
$Office365UserCredential = Get-Credential -UserName Administrator@Company.com.au -Message "Input Office365 Admin Account Creds"
$SpanningAdminEmail = "Administrator@Company.com.au"
$OutlookEmail = "Administrator@Company.com.au"
$exportlocation = "C:\scripts" #enter the path to your export here !NO TRAILING BACKSLASH!

$Tuser = Read-Host -Prompt "Input the name of the person (eg. Bob Ross)"
$confirm = Read-Host -prompt "Confirming the Name is $Tuser (Yes/No)"
$exporttemplate = @'
Container url: {ContainerURL*:https://xicnediscnam.blob.core.windows.net/da3fecb0-4ed4-447e-0315-08d5adad8a5a}; SAS token: {SASToken:?sv=2014-02-14&sr=c&si=eDiscoveryBlobPolicy9%7C0&sig=RACMSyH6Cf0k4EP2wZSoAa0QrhKaV38Oa9ciHv5Y8Mk%3D}; Scenario: General; Scope: BothIndexedAndUnindexedItems; Scope details: AllUnindexed; Max unindexed size: 0; File type exclusions for unindexed: <null>; Total sources: 2; Exchange item format: Msg; Exchange archive format: IndividualMessage; SharePoint archive format: SingleZip; Include SharePoint versions: True; Enable dedupe: EnableDedupe:True; Reference action: "<null>"; Region: ; Started sources: StartedSources:3; Succeeded sources: SucceededSources:1; Failed sources: 0; Total estimated bytes: 12,791,334,934; Total estimated items: 143,729; Total transferred bytes: {TotalTransferredBytes:7,706,378,435}; Total transferred items: {TotalTransferredItems:71,412}; Progress: {Progress:49.69%}; Completed time: ; Duration: {Duration:00:50:43.9321895}; Export status: {ExportStatus:DistributionCompleted}
Container url: {ContainerURL*:https://zgrbediscnam.blob.core.windows.net/5c21f7c7-42a2-4e24-9e69-08d5acf316f5}; SAS token: {SASToken:?sv=2014-02-14&sr=c&si=eDiscoveryBlobPolicy9%7C0&sig=F6ycaX5eWcRBCS1Z5nfoTKJWTrHkAciqbYRP5%2FhsUOo%3D}; Scenario: General; Scope: BothIndexedAndUnindexedItems; Scope details: AllUnindexed; Max unindexed size: 0; File type exclusions for unindexed: <null>; Total sources: 1; Exchange item format: FxStream; Exchange archive format: PerUserPst; SharePoint archive format: IndividualMessage; Include SharePoint versions: True; Enable dedupe: True; Reference action: "<null>"; Region: ; Started sources: 2; Succeeded sources: 2; Failed sources: 0; Total estimated bytes: 69,952,559,461; Total estimated items: 107,707; Total transferred bytes: {TotalTransferredBytes:70,847,990,489}; Total transferred items: {TotalTransferredItems:100,808}; Progress: {Progress:93.59%}; Completed time: 4/27/2018 11:45:46 PM; Duration: {Duration:04:31:21.1593737}; Export status: {ExportStatus:Completed}
'@

If ($confirm -eq "No") {
    Clear-Host
    Write-Host "Please close and re-run the script" -ForegroundColor RED
}



    #Connects to DC1 moves, removes groups, changes password of user

Write-Host "Input DC Admin credentials"
$SessionDC = New-PSSession -ComputerName "DC" -Credential $UserCredential -Name DC
Invoke-Command -Session $SessionDC -ArgumentList $Tuser -ScriptBlock {

        param($tuser)

    Import-Module ActiveDirectory

    $TuserObject = get-aduser -Filter 'Name -like $Tuser'
    $TOU = Get-ADObject -filter 'Name -like "Terminated Users"'
    $TUserGroups = Get-ADPrincipalGroupMembership ($TuserObject).ObjectGUID
    $Username = $TuserObject.SamAccountName

	    # Changes password to what ever is input

    set-adaccountpassword $TuserObject.ObjectGUID -reset -newpassword (Read-Host -AsSecureString "New Password")

    Write-Host "Password has been changed"

	    # Moves User to Terminated Users

    Move-ADObject ($TuserObject).ObjectGUID -TargetPath ($TOU).objectGUID

    Write-Host "User has been moved to Terminated Users"

	    # Removes user from all groups except Domain Users

    ForEach ($TuserGroup in $TUserGroups) {
        if ($TUserGroup.name -ne "Domain Users") {Remove-ADGroupMember -Identity $TuserGroup.name -Members $TUserObject.ObjectGUID -Confirm:$False}
    }

    Write-Host "Groups have been removed"
        
        # Moves DC2 Folder to Backup Server

    Copy-Item \\Server -Destination "\\Destination" -Recurse

    Write-Host "DC folders copied to \\Server"

    Read-Host -prompt "Check if files are moved to \\Server"

    Get-ChildItem \\Server\$Username -Recurse | Remove-Item -Force

        #check if someone has a computer in AD, if so asks if they want to move the computer to the standard Computer directory

    $ADComputerDir = Get-ADObject -filter 'cn -like "Computers"'
    $ADComputerObject = Get-ADComputer -Filter {Description -like $Tuser} | Select-Object -ExpandProperty ObjectGUID
    $ADComputerName = Get-ADComputer -Filter {Description -like $Tuser} | Select-Object -ExpandProperty Name
    
    if ($ADComputerName) {
        if ((Read-Host "The following computers are under $Tuser's Name. Would you like to move $ADComputerName to Default OU?") -eq "Yes") {
                foreach ($A in $ADComputerObject) {
                    move-adobject $A -TargetPath $ADComputerDir
            }
        }
    }
    $Username
}

Remove-PSSession -Session $SessionDC

$Lastday = read-host "Is today the $Tuser last day? (Yes/No)"
$TLastDay = if ($lastday -eq "Yes") {
    Write-Host "Inputting Today's Date"
    get-date -UFormat "%e/%m/%Y"
        } Else {
        Read-Host "Input Last Day dd/mm/yyyy"
    }

$TDate = (Get-Date).AddDays(30) | Get-Date -UFormat "%e/%m/%Y"
$CurrentDate = Get-Date -UFormat "%e/%m/%Y"

#region Mail autoreply and forward

    #Setup Mail redirection and autoreply

Write-Host "Input Office365 Admin Credentials"
$Session0365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Office365UserCredential -Authentication  Basic -AllowRedirection
$UserRedirect = Read-Host "Who are $Tuser's emails being redirected to? (Input full name eg. Michael Scott)"

Write-Host "Setting up email forwarding and AutoReply"

Import-PSSession $Session0365 -CommandName Get-mailbox, set-mailbox, set-mailboxautoreplyconfiguration -AllowClobber
$RedirectEmail = (Get-Mailbox $UserRedirect).primarySMTPaddress
$Reply = "I have left Company as of $TLastDay. <br> Please email any further requests to $RedirectEmail <br> Regards, <br> $Tuser"
Set-Mailbox -Identity "$tuser" -DeliverToMailboxAndForward $true -ForwardingSMTPAddress "$RedirectEmail" 
Set-MailboxAutoReplyConfiguration -Identity $tuser -AutoReplyState Enabled -ExternalMessage "$Reply" -InternalMessage "$Reply"

Remove-PSSession -Session $Session0365

#endregion


Write-Host "Setting up Outlook reminder to delete account in 30 days"

#region Create Outlook Reminder

$ol = New-Object -ComObject Outlook.Application
$meeting = $ol.CreateItem('olAppointmentItem')
$meeting.Subject = "Delete $Tuser Account"
$meeting.Body = 'Delete Account'
$meeting.Location = 'Virtual'
$meeting.ReminderSet = $true
$meeting.Importance = 1
$meeting.MeetingStatus = [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeeting
$meeting.Recipients.Add("$OutlookEmail")
$meeting.ReminderMinutesBeforeStart = 15
$meeting.Start = [datetime]::Today.Adddays(30).Addhours(9)
$meeting.Duration = 30
$meeting.Send()

#endregion

Write-Host "Made the reminder, and the email redirections"

#region Department Select
$DepartmentSelect = Read-Host "What Department are they in?"
#endregion

#region Word Document Creation
$Word = New-Object -ComObject Word.Application
$Word.Visible = $True
$Word.Documents.Add()
$Selection = $Word.Selection
$Selection.Font.Size = "14"
$Selection.TypeText("Staff Termination")
$Selection.TypeParagraph()
$Selection.TypeParagraph()
$Selection.Font.Size = "12"
$Selection.TypeText("Staff Details")
$Selection.TypeParagraph()
$Selection.font.Size = "11"
$Selection.Font.Bold = 1
$Selection.TypeText("Name: ")
$Selection.Font.Bold = 0
$Selection.TypeText("$Tuser")
$Selection.TypeParagraph()
$Selection.Font.Bold = 1
$Selection.TypeText("Department: ")
$Selection.Font.Bold = 0
$Selection.TypeText("$Department")
$Selection.TypeParagraph()
$Selection.Font.Bold = 1
$Selection.TypeText("Employment End Date: ")
$Selection.Font.Bold = 0
$Selection.TypeText("$TLastDay")
$Selection.TypeParagraph()

$Selection.Font.Size = "12"
$Selection.TypeText("Account")
$Selection.TypeParagraph()
$Selection.font.Size = "11"
$Selection.Font.Bold = 1
$Selection.TypeText("User Account Password Changed: ")
$Selection.Font.Bold = 0
$Selection.TypeText("Yes $CurrentDate")
$Selection.TypeParagraph()
$Selection.Font.Bold = 1
$Selection.TypeText("User Account Removed from Security and Dist Groups: ")
$Selection.Font.Bold = 0
$Selection.TypeText("Yes $CurrentDate")
$Selection.TypeParagraph()
$Selection.Font.Bold = 1
$Selection.TypeText("Remove VPN Access: ")
$Selection.Font.Bold = 0 
$Selection.TypeText("Yes $CurrentDate")
$Selection.TypeParagraph()
$Selection.Font.Bold = 1
$Selection.TypeText("Moved User to Terminated Users OU: ")
$Selection.Font.Bold = 0
$Selection.TypeText("Yes $CurrentDate")
$Selection.TypeParagraph()
$Selection.Font.Bold = 1
$Selection.TypeText("User Account Deletion Date ")
$Selection.Font.Bold = 0
$Selection.TypeText("Yes $TDate")
$Selection.TypeParagraph()

$Selection.Font.Size = "12"
$Selection.TypeText("Email")
$Selection.TypeParagraph()
$Selection.font.Size = "11"
$Selection.Font.Bold = 1
$Selection.TypeText("Email Redirection: ")
$Selection.Font.Bold = 0
$Selection.TypeText("Yes $CurrentDate $UserRedirect")
$Selection.TypeParagraph()
$Selection.Font.Bold = 1
$Selection.TypeText("Mailbox Access: ")
$Selection.Font.Bold = 0
$Selection.TypeText("Yes $CurrentDate")
$Selection.TypeParagraph()
$Selection.Font.Bold = 1
$Selection.TypeText("Out of Office Setup: ")
$Selection.Font.Bold = 0
$Selection.TypeText("Yes $CurrentDate")
$Selection.TypeParagraph()
$Selection.Font.Bold = 1
$Selection.TypeText("Email Archive: ")
$Selection.Font.Bold = 0
$Selection.TypeText("Yes $CurrentDate")
$Selection.TypeParagraph()

$Selection.Font.Size = "12"
$Selection.TypeText("Data")
$Selection.TypeParagraph()
$Selection.font.Size = "11"
$Selection.Font.Bold = 1
$Selection.TypeText("Backup Data on User's PC: ")
$Selection.Font.Bold = 0
$Selection.TypeText("Yes $CurrentDate")
$Selection.TypeParagraph()
$Selection.Font.Bold = 1
$Selection.TypeText("Backup Users Data on OneDrive: ")
$Selection.Font.Bold = 0
$Selection.TypeText("Yes $CurrentDate")
$Selection.TypeParagraph()
$Selection.Font.Bold = 1
$Selection.TypeText("Provide Access to Management: ")
$Selection.Font.Bold = 0
$Selection.TypeText("Yes $CurrentDate")
$Selection.TypeParagraph()
#endregion

    #Setup Mail Search
Write-Host "Preparing Mailbox to be downloaded"
$Session0365Search = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Office365UserCredential -Authentication  Basic -AllowRedirection
Import-PSSession $Session0365Search -CommandName *ComplianceSearch* -AllowClobber
New-ComplianceSearch -Name "$Tuser" -ExchangeLocation "$TUser" | Start-ComplianceSearch
Write-Host "Waiting for Search to Complete"
do
    {
        Start-Sleep -s 5
        $ComplianceSearch = get-complianceSearch "$Tuser"

    }
while ($ComplianceSearch.Status -ne 'Completed')

    #Export Mail Search
New-ComplianceSearchAction -SearchName $Tuser -Export -Format FxStream -ExchangeArchiveFormat PerUserPst
$TuserExport = $Tuser + "_Export"
Write-Host "Waiting for Export to Complete"
do
    {
        Start-Sleep -s 60
        $ComplianceExport = get-ComplianceSearchAction -Identity $TuserExport -IncludeCredential -Details | Select-Object -ExpandProperty Results | ConvertFrom-String -TemplateContent $exporttemplate
        $ComplianceExportProgress = $ComplianceExport.Progress
        Write-Host "Export still exporting, progress is $ComplianceExportProgress, waiting 60 seconds"
    }
while ($ComplianceExportProgress -ne '100.00%')

    #Download Exported Mailbox
    #Further assistance check out https://techcommunity.microsoft.com/t5/Office-365/Export-to-PST-via-Powershell/td-p/95007

$exportexe = "C:\UnifiedExportTool\microsoft.office.client.discovery.unifiedexporttool.exe" #path to your microsoft.office.client.discovery.unifiedexporttool.exe file. Usually found somewhere in %LOCALAPPDATA%\Apps\2.0\

    # Gather the URL and Token from the export in order to start the download

$exportdetails = Get-ComplianceSearchAction -Identity $TuserExport -IncludeCredential -Details | Select-Object -ExpandProperty Results | ConvertFrom-String -TemplateContent $exporttemplate
$exportdetails
$exportcontainerurl = $exportdetails.ContainerURL
$exportsastoken = $exportdetails.SASToken


    # Download the exported files from Office 365
Write-Host "Initiating download"
Write-Host "Saving export to: " $exportlocation
$arguments = "-name ""$Tuser""","-source ""$exportcontainerurl""","-key ""$exportsastoken""","-dest ""$exportlocation""","-trace true"
Start-Process -FilePath "$exportexe" -ArgumentList $arguments

    #Do while microsoft.office.client.discovery.unifiedexporttool.exe running
$started = $false
Do { $status = Get-Process microsoft.office.client.discovery.unifiedexporttool -ErrorAction SilentlyContinue
If (!($status)) {

Write-Host 'Waiting for process to start' ; Start-Sleep -Seconds 5 }

Else {

Write-Host 'Process has started' ; $started = $true

}

}Until ( $started )  

 

Do{

$ProcessesFound = Get-Process | Where-Object {$_.Name -like "*unifiedexporttool*"}
If ($ProcessesFound) { 

Write-Host "Export still downloading, waiting 60 seconds"
Start-Sleep -s 60

}

}Until (!$ProcessesFound)

    #Removes License from Spanning
$TuserEmail = Read-Host -Prompt "Input Username with @company.com.au eg Administrator@Company.com"
Import-Module SpanningO365
Get-SpanningAuthentication -ApiToken API -Region AP -AdminEmail $SpanningAdminEmail
Disable-SpanningUser -UserPrincipalName "$TuserEmail"

Import-Module JiraPS
Set-JiraConfigServer -Server "http://help.Company.com.au"
$creds = Get-Credential -UserName Administrator -Message "Jira account eg JIRAAdmin"
New-JiraSession -Credential $creds
$JiraIssue = Read-Host -Prompt "Input Jira Ticket Number"
$Ticket = "JIRA-" + $JiraIssue

$JiraComment = @"
$TUser's account has been terminated. Files have been moved to their backup folder.

*If something is incorrect, this is all automatic from account termination script.*
"@

$JiraWorkLog = @"
Account has been terminated by script saving upwards of 30 minutes per user terminated. 
"@

Add-JiraIssueComment -Issue $Ticket -Comment "$JiraComment"  
Add-JiraIssueWorklog -Issue $Ticket -TimeSpent "00:20" -Comment $JiraWorkLog -DateStarted (Get-Date)


Read-Host -prompt "$TUser has successfully been terminated. Please move their .pst any other files from their computer to the Backup Server. Press Enter to close."