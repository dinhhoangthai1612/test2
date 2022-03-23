Sub UpdateChart(LastColumnHide As String, LastColumnLetter As String, chartName1 As String, chartName2 As String)
    Dim sheetName As String
    Dim iRow, index As Integer
    sheetName = "Monthly_Chart"
    
    Worksheets(sheetName).Select
            
    'Chart 1
    Worksheets(sheetName).ChartObjects(chartName1).Activate
    'quet lai vung chon ngay cua bieu do
    For iRow = 1 To 4
        ActiveChart.SeriesCollection(iRow).XValues = "=" & sheetName & "!$" & LastColumnHide & "$26:$" & LastColumnLetter & "$26"
    Next iRow
    
    index = 27
    For iRow = 1 To 4
        ActiveChart.SeriesCollection(iRow).Values = "=" & sheetName & "!$" & LastColumnHide & "$" & index & ":$" & LastColumnLetter & "$" & index
        If index = 27 Or index = 30 Then
            index = index + 2
        ElseIf index = 29 Then
            index = index + 1
        End If
    Next iRow
    
    'Chart 2
    Worksheets(sheetName).ChartObjects(chartName2).Activate
    'quet lai vung chon ngay cua bieu do
    For iRow = 1 To 2
        ActiveChart.SeriesCollection(iRow).XValues = "=" & sheetName & "!$" & LastColumnHide & "$26:$" & LastColumnLetter & "$26"
    Next iRow
    
    index = 28
    For iRow = 1 To 2
        ActiveChart.SeriesCollection(iRow).Values = "=" & sheetName & "!$" & LastColumnHide & "$" & index & ":$" & LastColumnLetter & "$" & index
        index = index + 3
    Next iRow
End Sub