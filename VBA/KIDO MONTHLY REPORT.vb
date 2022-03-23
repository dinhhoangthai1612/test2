Sub UpdateChartSheetInputExFac(strFirstCell As String, strLastCell As String, indexDate As Integer, index As Integer, sheetName As String, chartName As String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Worksheets(sheetName).ChartObjects(chartName).Activate
    ActiveChart.ChartArea.Select
    
    For iRow = 1 To 4
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "=" & sheetName & "!$"& strFirstCell &"$" & indexDate & ":$" & strLastCell & "$" & indexDate & ""
    Next iRow
    
    For iRow = 1 To 4
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(iRow).Values = "=" & sheetName & "!$" & strFirstCell &"$" & index & ":$" & strLastCell & "$" & index
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    index = index + 1
    Next iRow
    
End Sub

Sub UpdateChartCoilVislonMetal(strFirstCell As String, strLastCell As String, indexDate As Integer, indexLeadTime As Integer, index As Integer, sheetName As String, chartName As String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Worksheets(sheetName).ChartObjects(chartName).Activate
    ActiveChart.ChartArea.Select
    
    For iRow = 1 To 5
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "=" & sheetName & "!$" & strFirstCell & "$" & indexDate & ":$" & strLastCell & "$" & indexLeadTime & ""
    Next iRow
    
    For iRow = 1 To 5
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(iRow).Values = "=" & sheetName & "!$" & strFirstCell & "$" & index & ":$" & strLastCell & "$" & index
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    index = index + 1
    Next iRow
    
End Sub

Sub HideColumn(sheetName As String, index As Integer, strLastCell As String)
    Sheets(sheetName).Range("B" & index & ":" & strLastCell).EntireColumn.Hidden = True
End Sub

Sub CheckFilter(sheetName as String)
    If Sheets(sheetName).FilterMode Then Sheets(sheetName).ShowAllData
End Sub

Sub UnhideColumn(sheetName As String)
    Worksheets(sheetName).Select
    Columns.EntireColumn.Hidden = False
End Sub

Sub SaveChartImage(sheetName As String, chartName As String, pathImage As String)
    Application.DisplayAlerts = False
    Dim i As Integer
    Dim ChtObj As ChartObject

    Worksheets(sheetName).Select
    For Each ChtObj In Worksheets(sheetName).ChartObjects
        If ChtObj.Name = chartName Then
            ChtObj.Activate
            ActiveChart.Export pathImage
        End If
        If ChtObj.Name = chartName Then Exit For
    Next ChtObj
End Sub

Sub AutoFitColumn(stColumn As String)
    Worksheets("Input_Ex_Fac").Select
    Worksheets("Input_Ex_Fac").Columns(stColumn).AutoFit
End Sub