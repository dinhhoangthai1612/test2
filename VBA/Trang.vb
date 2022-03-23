Sub drawBorderSampleLT(lastRow As String)
    Application.DisplayAlerts = False
    Worksheets("Report").Range(lastRow).Borders.LineStyle = xlContinuous
End Sub

Sub insertColumnSampleLT(columnName As String)
    Application.DisplayAlerts = False
    Worksheets("Report").Activate
    Columns(columnName & ":" & columnName).Select
    Selection.Insert Shift:=xlToRight
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

Sub addSheet()
Worksheets.Add().Name = "Trang"
End Sub





