Sub ChangeTitleChart(sheetName as String, chartName as String, chartTitle as string)
Sheets(sheetName).ChartObjects(chartName).Chart.ChartTitle.Text = chartTitle
End Sub

Sub FitColumn(get_Column as String, get_sheet as String)
Worksheets(get_sheet).Columns(get_Column & ":" & get_Column).AutoFit
End Sub