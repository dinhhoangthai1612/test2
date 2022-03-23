Sub drawBorderSampleLT(lastRow As String)
    Application.DisplayAlerts = False
    Worksheets("Report").Range(lastRow).Borders.LineStyle = xlContinuous
End Sub

Sub ShowAllData(sheetName As String)
    On Error Resume Next
    Worksheets(sheetName).ShowAllData
End Sub

Sub insertColumnSampleLT(columnName As String)
    Application.DisplayAlerts = False
    Worksheets("Report").Activate
    Columns(columnName & ":" & columnName).Select
    Selection.Insert Shift:=xlToRight
End Sub


Sub SortLargestToSmallest()
    Range("A2", Range("T" & Rows.Count).End(xlUp)).Sort [H2], xlDescending, Header:=xlYes
End Sub

Sub SaveChartImageSampleLT(sheetName As String, chartName As String, pathImage As String)
    Application.DisplayAlerts = False
    Dim i As Integer
    Dim oCht As ChartObject

    Sheets(sheetName).Activate
    For Each ChtObj In Worksheets(sheetName).ChartObjects
        If ChtObj.Name = chartName Then
            ChtObj.Activate
            ActiveChart.Export pathImage
        End If
        If ChtObj.Name = chartName Then Exit For
    Next ChtObj
End Sub

Sub SetChartDataDefectSample(sheetName As String, lastRow As String)
    Application.DisplayAlerts = False
    Worksheets(sheetName).ChartObjects("Chart 1").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).Values = "='" & sheetName & "'!$B$3:$" & lastRow & "3"
    ActiveChart.SeriesCollection(1).Name = "='" & sheetName & "'!$A$3"
    ActiveChart.SeriesCollection(1).XValues = "='" & sheetName & "'!$B$2:$" & lastRow & "2"
    ActiveChart.SeriesCollection(2).Values = "='" & sheetName & "'!$B$4:$" & lastRow & "4"
    ActiveChart.SeriesCollection(2).Name = "='" & sheetName & "'!$A$4"
    ActiveChart.SeriesCollection(3).Values = "='" & sheetName & "'!$B$5:$" & lastRow & "5"
    ActiveChart.SeriesCollection(3).Name = "='" & sheetName & "'!$A$5"
    ActiveChart.SeriesCollection(4).Values = "='" & sheetName & "'!$B$6:$" & lastRow & "6"
    ActiveChart.SeriesCollection(4).Name = "='" & sheetName & "'!$A$6"
    ActiveChart.SeriesCollection(5).Values = "='" & sheetName & "'!$B$8:$" & lastRow & "8"
    ActiveChart.SeriesCollection(5).Name = "='" & sheetName & "'!$A$8"
    ActiveChart.SeriesCollection(6).Values = "='" & sheetName & "'!$LX$13:$" & lastRow & "13"
    ActiveChart.SeriesCollection(6).Name = "='" & sheetName & "'!$A$13"
    ActiveChart.SeriesCollection(7).Values = "='" & sheetName & "'!$AKS$7:$" & lastRow & "7"
    ActiveChart.SeriesCollection(7).Name = "='" & sheetName & "'!$A$7"

    Worksheets(sheetName).ChartObjects("Chart 2").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).Values = "='" & sheetName & "'!$B$43:$" & lastRow & "43"
    ActiveChart.SeriesCollection(1).Name = "='" & sheetName & "'!$A$43"
    ActiveChart.SeriesCollection(1).XValues = "='" & sheetName & "'!$B$42:$" & lastRow & "42"
    ActiveChart.SeriesCollection(2).Values = "='" & sheetName & "'!$B$44:$" & lastRow & "44"
    ActiveChart.SeriesCollection(2).Name = "='" & sheetName & "'!$A$44"
    ActiveChart.SeriesCollection(3).Values = "='" & sheetName & "'!$B$45:$" & lastRow & "45"
    ActiveChart.SeriesCollection(3).Name = "='" & sheetName & "'!$A$45"
    ActiveChart.SeriesCollection(4).Values = "='" & sheetName & "'!$B$46:$" & lastRow & "46"
    ActiveChart.SeriesCollection(4).Name = "='" & sheetName & "'!$A$46"
    ActiveChart.SeriesCollection(5).Values = "='" & sheetName & "'!$B$47:$" & lastRow & "47"
    ActiveChart.SeriesCollection(5).Name = "='" & sheetName & "'!$A$47"

    Worksheets(sheetName).ChartObjects("Chart 3").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).Values = "='" & sheetName & "'!$B$50:$" & lastRow & "50"
    ActiveChart.SeriesCollection(1).Name = "='" & sheetName & "'!$A$50"
    ActiveChart.SeriesCollection(1).XValues = "='" & sheetName & "'!$B$49:$" & lastRow & "49"
    ActiveChart.SeriesCollection(2).Values = "='" & sheetName & "'!$B$51:$" & lastRow & "51"
    ActiveChart.SeriesCollection(2).Name = "='" & sheetName & "'!$A$51"
    ActiveChart.SeriesCollection(3).Values = "='" & sheetName & "'!$B$52:$" & lastRow & "52"
    ActiveChart.SeriesCollection(3).Name = "='" & sheetName & "'!$A$52"
    ActiveChart.SeriesCollection(4).Values = "='" & sheetName & "'!$B$53:$" & lastRow & "53"
    ActiveChart.SeriesCollection(4).Name = "='" & sheetName & "'!$A$53"

    Worksheets(sheetName).ChartObjects("Chart 4").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).Values = "='" & sheetName & "'!$B$56:$" & lastRow & "56"
    ActiveChart.SeriesCollection(1).Name = "='" & sheetName & "'!$A$56"
    ActiveChart.SeriesCollection(1).XValues = "='" & sheetName & "'!$B$55:$" & lastRow & "55"
    ActiveChart.SeriesCollection(2).Values = "='" & sheetName & "'!$B$57:$" & lastRow & "57"
    ActiveChart.SeriesCollection(2).Name = "='" & sheetName & "'!$A$57"
    ActiveChart.SeriesCollection(3).Values = "='" & sheetName & "'!$B$58:$" & lastRow & "58"
    ActiveChart.SeriesCollection(3).Name = "='" & sheetName & "'!$A$58"
    ActiveChart.SeriesCollection(4).Values = "='" & sheetName & "'!$B$59:$" & lastRow & "59"
    ActiveChart.SeriesCollection(4).Name = "='" & sheetName & "'!$A$59"

    Worksheets(sheetName).ChartObjects("Chart 5").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).Values = "='" & sheetName & "'!$B$62:$" & lastRow & "62"
    ActiveChart.SeriesCollection(1).Name = "='" & sheetName & "'!$A$62"
    ActiveChart.SeriesCollection(1).XValues = "='" & sheetName & "'!$B$61:$" & lastRow & "61"
    ActiveChart.SeriesCollection(2).Values = "='" & sheetName & "'!$B$63:$" & lastRow & "63"
    ActiveChart.SeriesCollection(2).Name = "='" & sheetName & "'!$A$63"
    ActiveChart.SeriesCollection(3).Values = "='" & sheetName & "'!$B$64:$" & lastRow & "64"
    ActiveChart.SeriesCollection(3).Name = "='" & sheetName & "'!$A$64"
    ActiveChart.SeriesCollection(4).Values = "='" & sheetName & "'!$B$65:$" & lastRow & "65"
    ActiveChart.SeriesCollection(4).Name = "='" & sheetName & "'!$A$65"

    Worksheets(sheetName).ChartObjects("Chart 6").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).Values = "='" & sheetName & "'!$B$68:$" & lastRow & "68"
    ActiveChart.SeriesCollection(1).Name = "='" & sheetName & "'!$A$68"
    ActiveChart.SeriesCollection(1).XValues = "='" & sheetName & "'!$B$67:$" & lastRow & "67"
    ActiveChart.SeriesCollection(2).Values = "='" & sheetName & "'!$B$69:$" & lastRow & "69"
    ActiveChart.SeriesCollection(2).Name = "='" & sheetName & "'!$A$69"
    ActiveChart.SeriesCollection(3).Values = "='" & sheetName & "'!$B$70:$" & lastRow & "70"
    ActiveChart.SeriesCollection(3).Name = "='" & sheetName & "'!$A$70"
    ActiveChart.SeriesCollection(4).Values = "='" & sheetName & "'!$B$71:$" & lastRow & "71"
    ActiveChart.SeriesCollection(4).Name = "='" & sheetName & "'!$A$71"

    Worksheets(sheetName).ChartObjects("Chart 11").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).Values = "='" & sheetName & "'!$B$193:$" & lastRow & "193"
    ActiveChart.SeriesCollection(1).Name = "='" & sheetName & "'!$A$193"
    ActiveChart.SeriesCollection(1).XValues = "='" & sheetName & "'!$B$192:$" & lastRow & "192"
    ActiveChart.SeriesCollection(2).Values = "='" & sheetName & "'!$B$194:$" & lastRow & "194"
    ActiveChart.SeriesCollection(2).Name = "='" & sheetName & "'!$A$194"
    ActiveChart.SeriesCollection(3).Values = "='" & sheetName & "'!$B$195:$" & lastRow & "195"
    ActiveChart.SeriesCollection(3).Name = "='" & sheetName & "'!$A$195"
    ActiveChart.SeriesCollection(4).Values = "='" & sheetName & "'!$B$196:$" & lastRow & "196"
    ActiveChart.SeriesCollection(4).Name = "='" & sheetName & "'!$A$196"
    ActiveChart.SeriesCollection(5).Values = "='" & sheetName & "'!$B$197:$" & lastRow & "197"
    ActiveChart.SeriesCollection(5).Name = "='" & sheetName & "'!$A$197"

    Worksheets(sheetName).ChartObjects("Chart 12").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).Values = "='" & sheetName & "'!$B$221:$" & lastRow & "221"
    ActiveChart.SeriesCollection(1).Name = "='" & sheetName & "'!$A$221"
    ActiveChart.SeriesCollection(1).XValues = "='" & sheetName & "'!$B$220:$" & lastRow & "220"
    ActiveChart.SeriesCollection(2).Values = "='" & sheetName & "'!$B$222:$" & lastRow & "222"
    ActiveChart.SeriesCollection(2).Name = "='" & sheetName & "'!$A$222"
    ActiveChart.SeriesCollection(3).Values = "='" & sheetName & "'!$B$223:$" & lastRow & "223"
    ActiveChart.SeriesCollection(3).Name = "='" & sheetName & "'!$A$223"
    ActiveChart.SeriesCollection(4).Values = "='" & sheetName & "'!$B$224:$" & lastRow & "224"
    ActiveChart.SeriesCollection(4).Name = "='" & sheetName & "'!$A$224"
    ActiveChart.SeriesCollection(5).Values = "='" & sheetName & "'!$B$225:$" & lastRow & "225"
    ActiveChart.SeriesCollection(5).Name = "='" & sheetName & "'!$A$225"

    Worksheets(sheetName).ChartObjects("Chart 15").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).Values = "='" & sheetName & "'!$B$254:$" & lastRow & "254"
    ActiveChart.SeriesCollection(1).Name = "='" & sheetName & "'!$A$254"
    ActiveChart.SeriesCollection(1).XValues = "='" & sheetName & "'!$B$253:$" & lastRow & "253"
    ActiveChart.SeriesCollection(2).Values = "='" & sheetName & "'!$B$255:$" & lastRow & "255"
    ActiveChart.SeriesCollection(2).Name = "='" & sheetName & "'!$A$255"
    ActiveChart.SeriesCollection(3).Values = "='" & sheetName & "'!$B$256:$" & lastRow & "256"
    ActiveChart.SeriesCollection(3).Name = "='" & sheetName & "'!$A$256"
    ActiveChart.SeriesCollection(4).Values = "='" & sheetName & "'!$B$257:$" & lastRow & "257"
    ActiveChart.SeriesCollection(4).Name = "='" & sheetName & "'!$A$257"
    ActiveChart.SeriesCollection(5).Values = "='" & sheetName & "'!$B$258:$" & lastRow & "258"
    ActiveChart.SeriesCollection(5).Name = "='" & sheetName & "'!$A$258"
    ActiveChart.SeriesCollection(6).Values = "='" & sheetName & "'!$B$259:$" & lastRow & "259"
    ActiveChart.SeriesCollection(6).Name = "='" & sheetName & "'!$A$259"
    ActiveChart.SeriesCollection(7).Values = "='" & sheetName & "'!$B$260:$" & lastRow & "260"
    ActiveChart.SeriesCollection(7).Name = "='" & sheetName & "'!$A$260"
End Sub

Sub FilterOnSheet(sheetname As String)
    Application.DisplayAlerts = True
    Sheets(sheetname).Activate
    Cells.AutoFilter
End Sub

Sub RemoveFilterOnSheet(sheetname As String)
    Application.DisplayAlerts = False
    Sheets(sheetname).Activate
    Cells.AutoFilter
End Sub

Sub RefreshOnlySheet(sheetName As String, rg as String)
    Application.DisplayAlerts = False
    With Sheets(sheetName).Activate
        Range(rg).Select
        Selection.ListObject.QueryTable.Refresh
        Application.CalculateUntilAsyncQueriesDone
    End With
End Sub

Sub WorkbookRefreshAll()
    Application.DisplayAlerts = False
    Application.CalculateFullRebuild
    ActiveWorkbook.RefreshAll
End Sub

Sub AllWorkbookPivots()
    Dim pt As PivotTable
    Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Worksheets
        For Each pt In ws.PivotTables
            If pt.Name <> "NoRefresh" Then
                pt.RefreshTable
            End If
        Next pt
    Next ws
End Sub





