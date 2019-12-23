Invoke-Command -ComputerName Server1 -Credential Administrator -ScriptBlock {
    $Names = Get-ChildItem C:\users\* | select -ExpandProperty Name
    foreach($N in $Names){
        $Path = "C:\Users\$N\Appdata\Local\Microsoft\Windows\Temporary Internet Files"
        if(Test-Path $Path){
            (Get-ChildItem $Path -Recurse | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue)
        }
        $Path2 = "C:\Users\$N\AppData\Local\Google\Chrome\User Data\Default\Cache"
        if(Test-Path $Path2){
            (Get-ChildItem $Path2 -Recurse | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue)
        }
    }
}