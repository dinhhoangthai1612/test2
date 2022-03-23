Sub UnhideColumn(sheetName as String)
    Sheets(sheetName ).Select
    Cells.Select
    Selection.EntireColumn.Hidden = False
End Sub

Sub RenameSheet(oldName as String, newName as String)
	
    Sheets(oldName).Select
    Sheets(oldName).Name = newName
End Sub

Sub InsertColumn(columnName as String)
    Columns(columnName).Select
    Selection.Insert Shift:=xlToLeft, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub

Sub WorkWithExcel(hideColumn as String, insertColumn as String,sourceColumn As String, destColumn As String, isHide as Boolean)
	Application.DisplayAlerts = False
	Application.CutCopyMode = False
	
	'Unhide rows
	'Cells.Select
    'Selection.EntireRow.Hidden = False
	
	'Insert column
	Columns(insertColumn).Select
    Selection.Insert Shift:=xlToLeft, CopyOrigin:=xlFormatFromLeftOrAbove
	
	'Auto fill range
    Columns(sourceColumn).Select
    Selection.AutoFill Destination:=Columns(destColumn), Type:=xlFillCopy
	
	'Copy value
	Columns(sourceColumn).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
		
	'Hide column
	if isHide = true then
	Columns(hideColumn).Select
    Selection.EntireColumn.Hidden = True
	end if
End Sub