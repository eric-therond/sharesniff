[System.Reflection.Assembly]::LoadFrom("C:\Program Files (x86)\Open XML SDK\V2.0\lib\DocumentFormat.OpenXml.dll") | out-null
[Reflection.Assembly]::LoadWithPartialName("DocumentFormat.OpenXml") | out-null
[Reflection.Assembly]::LoadWithPartialName("DocumentFormat.OpenXml.Packaging") | out-null
[Reflection.Assembly]::LoadWithPartialName("DocumentFormat.OpenXml.Spreadsheet") | out-null
[Reflection.Assembly]::LoadWithPartialName("OpenXmlPowerTools") | out-null
 
[DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]$Document = $null
$Document = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open("test5.xlsx", $false)
 $Document
[DocumentFormat.OpenXml.Packaging.WorkbookPart]$WorkBookPart = $Document.WorkbookPart
[DocumentFormat.OpenXml.Spreadsheet.Workbook]$WorkBook = $WorkBookPart.Workbook

"count sheets $($Workbook.Sheets.Count)"
foreach($sheet in $Workbook.Sheets) {
"here"
[DocumentFormat.OpenXml.Packaging.WorksheetPart]$workSheetPart = $workBookPart.GetPartById($sheet.Id)

$cells = Invoke-GenericMethod -InputObject $workSheetPart.Worksheet -MethodName Descendants -GenericType DocumentFormat.OpenXml.Spreadsheet.Cell

  foreach ($cell in $cells) {
    if ($cell.DataType.Value -eq "SharedString") {
      $stringTable = Invoke-GenericMethod -InputObject $workBookPart -MethodName GetPartsOfType -GenericType DocumentFormat.OpenXml.Packaging.SharedStringTablePart
      $value = $stringTable.SharedStringTable.InnerText
    } else {
      [String]$value = $cell.InnerText
    }

    "Value at {0}: {1}" -f $cell.CellReference, $value
    [DocumentFormat.OpenXml.Packaging.SharedStringTablePart]$sharedTablePart = $null
  }
  
}

$Document.Close()
