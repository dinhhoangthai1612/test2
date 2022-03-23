Sub FilterColumnSample()
    Sheets("Sheet1").Range("B1").AutoFilter Field:=2, Criteria1:="1"
End Sub

Sub CheckFilter(sheetName as String)
    Worksheets(sheetName).Select
    If Sheets(sheetName).FilterMode Then Sheets(sheetName).ShowAllData
End Sub

Sub ClearFillColor(index As String)
    Range("A" & index & ":C" & index).Interior.Color = xlNone
End Sub

Sub SetColor(index As String)
    Range("E" & index).Font.Color = RGB(0,112,192)
End Sub

Sub SetColorColumnE()
    Range("E2").Font.Color = RGB(0,0,0)
End Sub