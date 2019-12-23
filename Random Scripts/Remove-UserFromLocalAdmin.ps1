function Remove-UserFromLocalAdmin {
    [CmdletBinding()]
    param (
        [String]$Username,
        [String]$ComputerName,
        [PSCredential]$Credential
    )
    
    begin {
        $Arguments = @{
         class = "win32_process"
         name = "create"
         namespace = "root\cimv2"
         ArgumentList = 'powershell "remove-localgroupmember -group Administrators -member $Username"'
         }
    }
    
    process {
        Invoke-WmiMethod @Arguments -ComputerName $ComputerName -Credential $Credential
        }

    
    end {
        Write-Host "$Username has been removed from the local administrators group on $ComputerName"
    }
}