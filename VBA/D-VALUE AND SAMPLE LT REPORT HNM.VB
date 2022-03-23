Sub WorkbookRefreshQuery()
    Application.DisplayAlerts = False
    ActiveWorkbook.RefreshAll
End Sub

Sub SetChart(strCellEnd As String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Dim index As Integer
    Worksheets("Report").ChartObjects("Chart 2").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).XValues = "='Report'!$D$17:$" & strCellEnd & "$17"

    index = 1
    For iRow = 18 To 33
        ActiveChart.SeriesCollection(index).Values = "='Report'!$D$" & iRow & ":$" & strCellEnd & "$" & iRow
        ActiveChart.SeriesCollection(index).Name = "='Report'!$C$" & iRow
        index = index + 1
    Next iRow
End Sub

Sub RemoveFilterOnSheet(sheetname As String)
    Application.DisplayAlerts = False
    Sheets(sheetname).Activate
    Cells.AutoFilter
End Sub

Function RefreshOnlySheet(sheetname as String)
    Application.DisplayAlerts = False
    With Sheets(sheetname).Activate
        Range("A3").Select
        Selection.ListObject.QueryTable.Refresh
        Application.CalculateUntilAsyncQueriesDone
    End With
End Function

Function HideColumn(getColumn as String, columnName as String, getRange as String)
    
    Columns(getColumn).Hidden = True
    Application.Wait (Now + TimeValue("0:00:01"))

    Application.DisplayAlerts = False
    Worksheets("Report").Activate
    Columns(columnName & ":" & columnName).Select
    Selection.Insert Shift:=xlToRight

    Application.Wait (Now + TimeValue("0:00:01"))

    Dim rng As Range
    Set rng = Range(getRange)
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With

End Function

Sub SetChartSeriesDvalueReport(strCellValue as String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Dim index As Integer

    Worksheets("Report").ChartObjects("Chart 2").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(4).XValues = "='Report'!$F$5:$" & strCellValue & "$5"
    
    index = 1
    For iRow = 6 To 11
    ActiveChart.SeriesCollection(index).Values = "='Report'!$F$" & iRow & ":$" & strCellValue & "$" & iRow
    ActiveChart.SeriesCollection(index).Name = "='Report'!$E$" & iRow
    index = index + 1
    Next iRow
End Sub

Sub Sort_NoHeader()
Worksheets("Special cases").Activate
Range("A3", Range("C" & Rows.Count).End(xlUp)).Sort[B2], xlAscending
End Sub

Function HideColumn2(getColumn as String, columnName as String)
    
    Columns(getColumn).Hidden = True
    Application.Wait (Now + TimeValue("0:00:01"))

    Application.DisplayAlerts = False
    Worksheets("Report").Activate
    Columns(columnName & ":" & columnName).Select
    Selection.Insert Shift:=xlToRight

    Application.Wait (Now + TimeValue("0:00:01"))



End Function


Sub FillChart1(strCellValue as String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Dim index As Integer

    Worksheets("Report").ChartObjects("Chart 1").Activate
    ActiveChart.ChartArea.Select
    
    'CHART 5
    For iRow = 1 To 4
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "='Report'!$LK$2:$" & strCellValue & "$2"
    Next iRow

    index = 1
    For iRow = 3 To 7
    If iRow = 6 Then
    Else
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(index).Values = "='Report'!$C$" & iRow & ":$" & strCellValue & "$" & iRow
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    index = index + 1
    End If

    Next iRow
    'CHART 5
End Sub

Sub FillChart11(strCellValue As String)

    Application.DisplayAlerts = False
    Dim iRow As Integer
    Dim index As Integer

    Worksheets("Report").ChartObjects("Chart 11").Activate
    ActiveChart.ChartArea.Select
    
    'CHART 5
    For iRow = 1 To 5
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "='Report'!$EN$26:$" & strCellValue & "$26"
    Next iRow

    index = 1
    For iRow = 27 To 31
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(index).Values = "='Report'!$EN$" & iRow & ":$" & strCellValue & "$" & iRow
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    index = index + 1
    Next iRow
    'CHART 5
End Sub

Sub FillChart12(strCellValue AS String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Dim index As Integer

    Worksheets("Report").ChartObjects("Chart 12").Activate
    ActiveChart.ChartArea.Select
    
    'CHART 5
    For iRow = 1 To 5
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "='Report'!$EN$51:$" & strCellValue & "$51"
    Next iRow

    index = 1
    For iRow = 52 To 56
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(index).Values = "='Report'!$EN$" & iRow & ":$" & strCellValue & "$" & iRow
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    index = index + 1
    Next iRow
    'CHART 5
End Sub

Sub FillChart15(strCellValue As String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Dim index As Integer

    Worksheets("Report").ChartObjects("Chart 15").Activate
    ActiveChart.ChartArea.Select
    
    'CHART 5
    For iRow = 1 To 6
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "='Report'!$EP$73:$" & strCellValue & "$73"
    Next iRow

    index = 1
    For iRow = 74 To 80
    If iRow = 79 Then
    Else
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(index).Values = "='Report'!$JF$" & iRow & ":$" & strCellValue & "$" & iRow
    ActiveChart.SeriesCollection(index).Name = "='Report'!$A$" & iRow
    index = index + 1
    End If

    Next iRow
    'CHART 5
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