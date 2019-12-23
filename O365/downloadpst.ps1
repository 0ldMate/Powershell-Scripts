$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

Import-Module (Import-PSSession -Session $Session -AllowClobber -DisableNameChecking) -Global
$searchname = "Username" #enter your search name here
$exportlocation = "C:\exports" #enter the path to your export here !NO TRAILING BACKSLASH!
$exportexe = "C:\UnifiedExportTool\microsoft.office.client.discovery.unifiedexporttool.exe" #path to your microsoft.office.client.discovery.unifiedexporttool.exe file. Usually found somewhere in %LOCALAPPDATA%\Apps\2.0\

# Gather the URL and Token from the export in order to start the download
#We only need the ContainerURL and SAS Token but I parsed some other fields as well while working with AzCopy
#The Container URL and Token in the following template has been altered to protect the innocent:
$exporttemplate = @'
Container url: {ContainerURL*:https://xicnediscnam.blob.core.windows.net/da3fecb0-4ed4-447e-0315-08d5adad8a5a}; SAS token: {SASToken:?sv=2014-02-14&sr=c&si=eDiscoveryBlobPolicy9%7C0&sig=RACMSyH6Cf0k4EP2wZSoAa0QrhKaV38Oa9ciHv5Y8Mk%3D}; Scenario: General; Scope: BothIndexedAndUnindexedItems; Scope details: AllUnindexed; Max unindexed size: 0; File type exclusions for unindexed: <null>; Total sources: 2; Exchange item format: Msg; Exchange archive format: IndividualMessage; SharePoint archive format: SingleZip; Include SharePoint versions: True; Enable dedupe: EnableDedupe:True; Reference action: "<null>"; Region: ; Started sources: StartedSources:3; Succeeded sources: SucceededSources:1; Failed sources: 0; Total estimated bytes: 12,791,334,934; Total estimated items: 143,729; Total transferred bytes: {TotalTransferredBytes:7,706,378,435}; Total transferred items: {TotalTransferredItems:71,412}; Progress: {Progress:49.69%}; Completed time: ; Duration: {Duration:00:50:43.9321895}; Export status: {ExportStatus:DistributionCompleted}
Container url: {ContainerURL*:https://zgrbediscnam.blob.core.windows.net/5c21f7c7-42a2-4e24-9e69-08d5acf316f5}; SAS token: {SASToken:?sv=2014-02-14&sr=c&si=eDiscoveryBlobPolicy9%7C0&sig=F6ycaX5eWcRBCS1Z5nfoTKJWTrHkAciqbYRP5%2FhsUOo%3D}; Scenario: General; Scope: BothIndexedAndUnindexedItems; Scope details: AllUnindexed; Max unindexed size: 0; File type exclusions for unindexed: <null>; Total sources: 1; Exchange item format: FxStream; Exchange archive format: PerUserPst; SharePoint archive format: IndividualMessage; Include SharePoint versions: True; Enable dedupe: True; Reference action: "<null>"; Region: ; Started sources: 2; Succeeded sources: 2; Failed sources: 0; Total estimated bytes: 69,952,559,461; Total estimated items: 107,707; Total transferred bytes: {TotalTransferredBytes:70,847,990,489}; Total transferred items: {TotalTransferredItems:100,808}; Progress: {Progress:93.59%}; Completed time: 4/27/2018 11:45:46 PM; Duration: 04:31:21.1593737; Export status: {ExportStatus:Completed}
'@
$exportname = $searchname + "_Export"
$exportdetails = Get-ComplianceSearchAction -Identity $exportname -IncludeCredential -Details | select -ExpandProperty Results | ConvertFrom-String -TemplateContent $exporttemplate
$exportdetails
$exportcontainerurl = $exportdetails.ContainerURL
$exportsastoken = $exportdetails.SASToken


# Download the exported files from Office 365
Write-Host "Initiating download"
Write-Host "Saving export to: " $exportlocation
$arguments = "-name ""$searchname""","-source ""$exportcontainerurl""","-key ""$exportsastoken""","-dest ""$exportlocation""","-trace true"
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

$ProcessesFound = Get-Process | ? {$_.Name -like "*unifiedexporttool*"}
If ($ProcessesFound) { 

$Progress = get-ComplianceSearchAction -Identity $exportname -IncludeCredential -Details | select -ExpandProperty Results | ConvertFrom-String -TemplateContent $exporttemplate | %{$_.Progress}
Write-Host "Export still downloading, waiting 60 seconds"
Start-Sleep -s 60

}

}Until (!$ProcessesFound)

Remove-PSSession -Session $Session