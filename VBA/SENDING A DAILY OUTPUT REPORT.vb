Sub CheckHideColumnOnSheet()
    Dim lastColumn As Integer
    lastColumn = 50
    For i = 1 To lastColumn
        If Columns(i).EntireColumn.Hidden = True Then
            Columns(i).EntireColumn.Hidden = False
        End If
    Next
End Sub

Sub HideColumn(sheetName As String, strColumnHide As String)
    Sheets(sheetName).Range(strColumnHide).EntireColumn.Hidden = True
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