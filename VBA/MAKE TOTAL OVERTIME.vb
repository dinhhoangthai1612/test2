Sub RoundNumberInRange(setRange As String, setSheet As String)
    Dim getRow As Integer
    Dim getColumn As Integer
    Dim getValue As String
    
    Sheets(setSheet).Select
    For Each c In Range(setRange)
        getRow = c.Row
        getColumn = c.Column
        getValue = Trim(CStr(c.Value))
        
        If getValue <> "0" And Len(getValue) > 0 Then
            If InStr(1, c.Value, "1.8") = 1 Then
                Cells(getRow, getColumn) = "2"
            ElseIf InStr(1, c.Value, "3.8") = 1 Then
                Cells(getRow, getColumn) = "4"
            ElseIf InStr(1, c.Value, "11.8") = 1 Then
                Cells(getRow, getColumn) = "12"
            End If
        End If
    Next c
End Sub

Sub CheckFilter(sheetName as String)
    Worksheets(sheetName).Select
    If Sheets(sheetName).FilterMode Then Sheets(sheetName).ShowAllData
End Sub

Sub CheckHideColumnOnSheet(sheetName As String)
    Worksheets(sheetName).Select
    For Each cell In Range("B7:AR7").Cells
        If cell.EntireColumn.Hidden = True Then
            Columns(cell.Column).EntireColumn.Hidden = False
        End If
    Next cell
End Sub

Sub CopyNewSheet(sheetNameCopy As String, renameSheet As String)
    Sheets(sheetNameCopy).Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = renameSheet
End Sub

Sub SortColumnTOTAL(sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Worksheets(sheetName).Select
    ws.Range("AR6").EntireRow.Hidden = True
    ws.Range("AR7").Sort Key1:=Range("AR7"), Order1:=xlDescending, Header:=xlNo
End Sub