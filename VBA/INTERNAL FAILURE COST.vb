Sub Convert_Column()
    Range("A1:A1048576").NumberFormat = "@"
    Range("C1:C1048576").NumberFormat = "@"
    Range("D1:D1048576").NumberFormat = "@"
    Range("E1:E1048576").NumberFormat = "@"
End Sub

Sub DeleteData()
    Dim ws As Worksheet
    Dim lRow As Integer, lRowWs As Integer
    Dim yYear As String
    Dim yYearDelete As String
    
    Set ws = ThisWorkbook.Sheets("PASTE")
    lRowWs = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' delete current sort option if any
    If ws.FilterMode = True Then
        ws.ShowAllData
    End If
    
    ' sort data
    ws.Range("A4:AN" & lRowWs).Sort key1:=ws.Range("B4"), order1:=xlAscending, Header:=xlYes
    
    ' check the first year and delete rows
    yYear = Left(ws.Range("B5").Value2, 4)
    yYearDelete = Year(Now()) - 4
    If yYear = Year(Now()) - 5 Then
        yYearDelete = yYearDelete + "03"
        lRow = ws.Range("B:B").Find(what:=yYearDelete, lookat:=xlPart, searchdirection:=xlPrevious).Row
        If lRow >= 5 Then
            ws.Rows("5:" & lRow).Delete
        End If
    End If
End Sub

Sub CheckFilter(sheetName as String)
    Worksheets(sheetName).Select
    If Sheets(sheetName).FilterMode Then Sheets(sheetName).ShowAllData
End Sub

Sub Convert_Range(index As Integer, sheetName as String)
    Worksheets(sheetName).Select
    Range("B" & index & ":E" & index).NumberFormat = "@"
End Sub

Sub UnhideColumn(sheetName As String)
    Worksheets(sheetName).Select
    Columns.EntireColumn.Hidden = False
End Sub