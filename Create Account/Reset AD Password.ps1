$SAMAccountName = Read-Host -Prompt "Username to be reset"
$Manager = Read-Host -Prompt "Who is the Manager the password is being sent to?"
$ManagerEmail = Read-Host -Prompt "What is the managers email address?"
$UserCredential = Get-Credential -UserName Administrator -Message "Administrator"
$SessionDC = New-PSSession -ComputerName "DC" -Credential $UserCredential -Name DC
Import-Module -PSSession $SessionDC -Name ActiveDirectory
function New-PassPhrase {
    <#
    .SYNOPSIS
    Generate PassPhrase for account logins

    .DESCRIPTION
    Generate a PassPhrase from a pre-defined list of words instead of using random character passwords
    Inspiration https://millerb.co.uk/2018/08/18/Generating-Passphrases-Instead-Of-Passwords.html
    Inspiration https://github.com/RickFlist/PoSh/blob/master/Modules/MTL-PasswordGenerator/MTL-PasswordGenerator.psm1

    .PARAMETER MinLength
    Length of PassPhrase to be generated

    .PARAMETER Delimiter
    The Delimiter to be used when outputting the PassPhrase. If no delimiter is specified then a hyphen is used '-'

    .Parameter PhraseFile
    Path to a phrase file to use for the generation of passwords

    .EXAMPLE
    New-PassPhrase -MinLength 25

    .EXAMPLE
    New-PassPhrase -MinLength 25 -Delimiter ';'

    .NOTES
    NCSC UK Guidance on Secure Passwords
    https://www.ncsc.gov.uk/guidance/password-guidance-simplifying-your-approach
    #>
    [CmdletBinding()]
    param (
        [Parameter(Position = 1)]
        [int] $MinLength=15,

        [Parameter(Position = 2)]
        [char[]] $Delimiter = '-',

        [Parameter(Position = 3)]
        [string]$PhraseFile = "$PSScriptRoot\provide-your-own-wordlist.txt"
    )

    begin {
        if (Test-Path $PhraseFile) {
            $wordlist = ([String[]]@(Get-Content -Path $PhraseFile))
        } else {
            Write-Error "Phrase file count not be found"
            exit
        }
    }

    process {
        $phrasearr = @()
        while ($phrase.length -lt $MinLength) {
            $phrasearr += $wordlist | Get-Random -Count ($MinLength / 5)
            $phrase = $phrasearr -join $Delimiter
        }
    }

    end {
        $phrasearr -join $Delimiter
    }
}

$Name = Get-ADUser -Identity $SAMAccountName | Select-Object -ExpandProperty Name
$Pass = New-PassPhrase -MinLength 10
$Pass = $pass + (Get-Random -Maximum 200)
Set-ADAccountPassword $SAMAccountName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$pass" -force)
Set-ADUser -Identity $SAMAccountName -ChangePasswordAtLogon $true

$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)
$Mail.To = $ManagerEmail
$Mail.Subject = "$Name Account Information"
$Mail.Body = @"
Hi $Manager,
This is the password for $Name
Their Username is $SAMAccountName
Their password is $Pass
Once signing in they will be requested to change their password. Their new password must be 8 characters long, have a capital letter, a number and a symbol.
Thanks
Admin

*THIS IS AN AUTOMATED EMAIL FROM A SCRIPT. IF SOMETHING APPEARS INCORRECT PLEASE EMAIL ADMINISTRATOR.*
"@
$Mail.Send()

write-host "$Name's password has been reset and emailed to $ManagerEmail and will be reset at logon"

Import-Module JiraPS
$creds = Get-Credential -UserName jheadrick -Message "Jira account eg Administrator"
$TicketNumber = "JIRA-" + (Read-Host -Prompt "Jira Ticket Number (eg. 231)")
New-JiraSession -Credential $creds
$JiraComment = @"
$Name's account has been created. 
Email has been set to $Manager including the username, email, password and instructions on how to reset their password.
Password will be forced to change at login.

*This is an automated message*
"@

Add-JiraIssueComment -Issue "$TicketNumber" -Comment "$JiraComment"  

Read-Host -Prompt "$TicketNumber has been edited with this info"

$TuserEmail = $SAMAccountName + "@Company.com.au"
$SpanningAdminEmail = "Administrator@Company.com.au"
Import-Module SpanningO365
Get-SpanningAuthentication -ApiToken API -Region AP -AdminEmail $SpanningAdminEmail
Enable-SpanningUser -UserPrincipalName "$TuserEmail"

Write-Host "Spanning has been enabled on the user"