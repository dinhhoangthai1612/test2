Sub FillChartDye(strFirstCell As String, strLastCell As String, indexDate As Integer, index As Integer, sheetName As String, chartName As String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Worksheets(sheetName).ChartObjects(chartName).Activate
    ActiveChart.ChartArea.Select
    
    For iRow = 1 To 2
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "=" & sheetName & "!$"& strFirstCell &"$" & indexDate & ":$" & strLastCell & "$" & indexDate & ""
    Next iRow
    
    For iRow = 1 To 2
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(iRow).Values = "=" & sheetName & "!$" & strFirstCell &"$" & index & ":$" & strLastCell & "$" & index
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    index = index + 1
    Next iRow
    
End Sub

Sub ChangeTitleChartDye(sheetName As String, chartName As String, datetimeChart as String, strValue as String)
    Sheets(sheetName).ChartObjects(chartName).Chart.ChartTitle.Text = "DYED STOCK IN RECORD " & datetimeChart & " " & strValue
End Sub

Sub FillChartMPV(strFirstCell As String, strLastCell As String, indexDate As Integer, index As Integer, sheetName As String, chartName As String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Worksheets(sheetName).ChartObjects(chartName).Activate
    ActiveChart.ChartArea.Select
    
    For iRow = 1 To 7
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "=" & sheetName & "!$"& strFirstCell &"$" & indexDate & ":$" & strLastCell & "$" & indexDate & ""
    Next iRow
    
    For iRow = 1 To 7
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(iRow).Values = "=" & sheetName & "!$" & strFirstCell &"$" & index & ":$" & strLastCell & "$" & index
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    index = index + 1
    Next iRow
    
End Sub

Sub ChangeTitleChartMPV(sheetName As String, chartName As String, datetimeChart as String, strValue as String)
    Sheets(sheetName).ChartObjects(chartName).Chart.ChartTitle.Text = "STOCK OUT TO MPV " & datetimeChart & " " & strValue
End Sub