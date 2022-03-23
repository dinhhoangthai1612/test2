Sub deleteRow(sheetName As String, deleteRow As String)
    Sheets(sheetName).Select
    Rows(deleteRow).Select
	Selection.Delete Shift:=xlUp
End Sub

Sub ConvertToText()
    Cells.Select
    Selection.NumberFormat = "@"
    ActiveWorkbook.Save
End Sub