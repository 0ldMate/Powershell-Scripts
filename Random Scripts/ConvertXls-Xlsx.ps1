# Not written by me, can't remember the author
$xlWorkbookDefault = 51
$xlFixedFormat = $xlWorkbookDefault

$excel = New-Object -ComObject excel.application
$excel.visible = $false
$folderpath = "Path\*"
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