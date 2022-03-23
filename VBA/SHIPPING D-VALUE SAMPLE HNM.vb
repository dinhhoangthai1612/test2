Sub InsertColumn(getColumn as String)
    Application.DisplayAlerts = False
    Worksheets("4.3dataPR").Activate
    Columns(getColumn & ":" & getColumn).Select
    Selection.Insert Shift:=xlToRight
End Sub

Sub SetRangeBorder(getColumn as String, sheetName as String)
    Dim rng As Range

    Set rng = Range(getColumn & "2:" & getColumn & "111")

    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
End Sub

Sub FillChart(strFirstCell As String, strLastCell As String, indexDate As Integer, indexActAndFree As Integer, index As Integer, sheetName As String, chartName As String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Worksheets(sheetName).ChartObjects(chartName).Activate
    ActiveChart.ChartArea.Select
    
    For iRow = 1 To 5
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "=" & sheetName & "!$" & strFirstCell & "$" & indexDate & ":$" & strLastCell & "$" & indexActAndFree & ""
    Next iRow
    
    For iRow = 1 To 5
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(iRow).Values = "=" & sheetName & "!$" & strFirstCell & "$" & index & ":$" & strLastCell & "$" & index
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    index = index + 1
    Next iRow
    
End Sub

Sub HideColumn(sheetName As String, strFirstCell As String, index As Integer, strLastCell As String)
    Sheets(sheetName).Range(strFirstCell & index & ":" & strLastCell).EntireColumn.Hidden = True
End Sub

Sub SaveChartImage(sheetName As String, chartName As String, pathImage As String)
    Application.DisplayAlerts = False
    Dim i As Integer
    Dim ChtObj As ChartObject

    Sheets(sheetName).Activate
    For Each ChtObj In Worksheets(sheetName).ChartObjects
        If ChtObj.Name = chartName Then
            ChtObj.Activate
            ActiveChart.Export pathImage
        End If
        If ChtObj.Name = chartName Then Exit For
    Next ChtObj
End Sub

Sub ChangeTitleChart(datetimeChart As String)
    Sheets("4.4GRAPH").ChartObjects("Chart 12").Chart.ChartTitle.Text = "SHIPPING SAMPLE D_VALUE (" & datetimeChart & ")"
End Sub

Sub ChangeTitleChartDelay(datetimeChart as String)
    Sheets("4.5Delay").ChartObjects("Chart 1").Chart.ChartTitle.Text = "DELAY - " & datetimeChart
End Sub