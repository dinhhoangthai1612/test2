Sub RoundNumberInRange(setRange as String, setSheet as String)
Dim getRow As Integer
Dim getColumn As Integer
Dim getValue As String

Sheets(setSheet).Select
For Each c In Range(setRange)
	getRow = C.Row
    getColumn = C.Column
    getValue = Trim(CStr(C.Value))
    
    If getValue <> "0" And Len(getValue) > 0 Then
        If InStr(1, C.Value, "3.8") = 1 Then
            Cells(getRow, getColumn) = "4"
        ElseIf InStr(1, C.Value, "11.8") = 1 Then
            Cells(getRow, getColumn) = "12"
        End If
    End If
Next c
End Sub

Sub ConvertToText(rangeColumn As String)
    Columns(rangeColumn).Select
    Selection.NumberFormat = "@"
End Sub