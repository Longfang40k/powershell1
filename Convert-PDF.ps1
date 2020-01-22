$path = "C:\Users\jgathercole\Documents\test"
$destinationPath =  "C:\Users\jgathercole\Documents\test\pdf"
$xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type] 
$filter = Get-ChildItem -Path $path -include *.xls, *.xlsx -recurse 
$objExcel = New-Object -ComObject excel.application 
$objExcel.visible = $false 
foreach($wb in $filter) 
{ 
    $filepath = join-path -Path $destinationPath -ChildPath ($wb.BaseName + ".pdf")
    $workbook = $objExcel.workbooks.open($wb.fullname, 3) 
    $workbook.Saved = $true 
    "saving $filepath" 
    $workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $filepath) 
    $objExcel.Workbooks.close() 
} 
$objExcel.Quit()