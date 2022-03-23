Sub CheckFilterOnSheet(sheetName As String)
   If Sheets(sheetName).FilterMode Then Sheets(sheetName).ShowAllData
End Sub

Sub InsertMultipleRows(countRow As Integer)
    Worksheets("Form").Range("A18:I" & countRow).EntireRow.Insert
End Sub

Sub ConvertFormatCell(cellValue As String)
    Range(cellValue).NumberFormat = "@"
End Sub