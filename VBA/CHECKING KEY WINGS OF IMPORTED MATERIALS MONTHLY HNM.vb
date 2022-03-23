Sub FilterByDate(firstDate As String, endDate As String)
    'Format datetime as M/d/yyyy
    On Error GoTo 0
    ActiveSheet.Range("A3:AE1048576").AutoFilter Field:=16, Criteria1:= _
        ">=" & firstDate, Operator:=xlAnd, Criteria2:="<=" & endDate
End Sub

Sub ClearFilter()
	Application.DisplayAlerts = False
	
	If ActiveSheet.FilterMode = True Then
		ActiveSheet.ShowAllData
	End If
	
    Cells.Select
    Selection.EntireColumn.Hidden = False
    Selection.EntireRow.Hidden = False
	
    Application.DisplayAlerts = True
End Sub

Sub FormatHeader()
    Range("U3:W3").Select
    With Selection
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Sub AutoFitColumn()
    Cells.Select
    Selection.Columns.AutoFit
End Sub
