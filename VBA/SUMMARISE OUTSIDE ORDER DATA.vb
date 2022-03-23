Sub UnhideRowsAndColumns(sheetName AS String)
    Sheets(sheetName).Select
	Cells.Select
    Selection.EntireColumn.Hidden = False
    Selection.EntireRow.Hidden = False
    If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.AutoFilterMode = False
    End If
End Sub

Sub RenameSheet(sheetName As String)
    Sheets("Order yyyy").Select
    Sheets("Order yyyy").Name = sheetName
End Sub
