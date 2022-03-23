Sub DeleteColumn(columnName As String)
    Columns(columnName).Select
    Selection.Delete Shift:=xlToLeft
    Cells.Select
    Selection.Columns.AutoFit
End Sub

Sub SortByRange()
    Rows("20:156").Select
    Selection.AutoFilter
    With ActiveWorkbook.Worksheets("Assembly POP").AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("B21:B221"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .Apply
    End With
    Selection.AutoFilter
End Sub

Sub updateChart(sheetName As String, lastCell As String, chartName As String)
    ActiveSheet.ChartObjects(chartName).Activate
    ActiveChart.FullSeriesCollection(1).Select
    Selection.Formula = "=SERIES('" & sheetName & "'!$A$2,'" & sheetName & "'!$C$3:$" & lastCell & "$3,'" & sheetName & "'!$C$2:$" & lastCell & "$2,1)"
End Sub