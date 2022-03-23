Sub MakeCleanExcel(sheetName AS String)
	Application.DisplayAlerts = False
    Sheets(sheetName).Select
	Cells.Select
    Selection.EntireColumn.Hidden = False
    Selection.EntireRow.Hidden = False
    If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.AutoFilterMode = False
    End If
End Sub