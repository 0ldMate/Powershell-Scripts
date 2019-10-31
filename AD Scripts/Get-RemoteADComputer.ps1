function Get-RemoteADComputer {
    [CmdletBinding()]
    param (
        [String]$Name,
        [PSCredential]$Credential
    )
    
    begin {
        $DC = New-PSSession -ComputerName DC -Credential $Credential
    }
    
    process {
        Invoke-Command -Session $DC -ArgumentList $Name -ScriptBlock{
            param($Name)
        Import-Module -Name ActiveDirectory
        Get-ADComputer -Filter "description -like '$Name'" -Properties Name, Description
        }
    }
    
    end {
        Remove-PSSession -Session $DC
    }
}