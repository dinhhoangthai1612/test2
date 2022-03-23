Sub NewHideRow(index As Integer)
    Dim yYear As String
    Dim mMonth As String
    Dim dateValue As String
    Dim i As Integer
    
    i = index
    yYear = Year(Now()) - 1
    mMonth = Month(Now()) - 2
    If mMonth < 10 Then
        mMonth = "0" + mMonth
    End If
    dateValue = yYear + mMonth
    
    Sheets("Chart").Range("A11:A10000").EntireRow.Hidden = False
    For Each cell In Range("A11:A" & i)
        If cell.Value <= dateValue Then
            cell.EntireRow.Hidden = True
        End If 
    Next cell
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

Sub FillChart(strFirstCell As String, index1 As Integer, index2 As Integer, sheetName As String, chartName As String)
    Application.DisplayAlerts = False
    Dim iCol As Integer
    Dim arrayStr As Variant
    Dim countColumn As Integer
    arrayStr = Array("B", "C", "D", "E", "F", "G", "H", "I")
    countColumn = 0
    Worksheets(sheetName).ChartObjects(chartName).Activate
    ActiveChart.ChartArea.Select
    
    For iCol = 1 To 8
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iCol).XValues = "=" & sheetName & "!$" & strFirstCell & "$" & index1 & ":$" & strFirstCell & "$" & index2 & ""
    Next iCol
    
    For iCol = 1 To 8
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(iCol).Values = "=" & sheetName & "!$" & arrayStr(countColumn) & "$" & index1 & ":$" & arrayStr(countColumn) & "$" & index2
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    countColumn = countColumn + 1
    Next iCol
    
End Sub