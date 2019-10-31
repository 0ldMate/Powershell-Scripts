$xlWorkbookDefault = 51
$xlFixedFormat = $xlWorkbookDefault

$excel = New-Object -ComObject excel.application
$excel.visible = $false
$folderpath = "C:\Users\jheadrick\Mechellet\MMD Excel _FG Overseas_Electronic\*"
$filetype ="*xls"
Get-ChildItem -Path $folderpath -Include $filetype | 
ForEach-Object `
{
$path = ($_.fullname).substring(0,($_.FullName).lastindexOf("."))
"Converting $path to $filetype..."
$workbook = $excel.workbooks.open($_.fullname)

$workbook.saveas($path, $xlFixedFormat)
$workbook.close()
}
$excel.Quit()
$excel = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()