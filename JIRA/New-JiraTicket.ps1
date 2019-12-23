$JiraCredential = Get-Credential -Message "Jira Login" -UserName User
Import-Module JiraPS
Set-JiraConfigServer -Server "http://Jira.Company.com.au"
New-JiraSession -Credential $JiraCredential
Import-Csv "$PSScriptRoot\JiraTickets.csv" | ForEach-Object {
$Description = @"
Description
"@
$Reporter = "Reporter"
$Assignee = "Assigned user"
$Summary = $_.Summary
$Project = "Project"
$IssueType = "Marketing"
$fields = @{
    duedate = "2019-06-30"
}
New-JiraIssue -Project $Project -IssueType $IssueType -Summary $Summary -Description $Description -Reporter $Reporter -Fields $fields | Set-JiraIssue -Assignee $Assignee -PassThru -SkipNotification | Invoke-JiraIssueTransition -Transition 421

}