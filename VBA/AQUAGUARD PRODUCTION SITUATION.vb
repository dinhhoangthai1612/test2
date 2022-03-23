Sub WorkbookRefreshAll()
    Application.DisplayAlerts = False
    Application.CalculateFullRebuild
    ActiveWorkbook.RefreshAll
End Sub

Sub FillChart_graph(strFirstCell As String, strLastCell As String, indexDate As Integer, index As Integer, sheetName As String, chartName As String)
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

Sub FillChart_Sheet1(strFirstCell As String, strLastCell As String, indexDate As Integer, index As Integer, sheetName As String, chartName As String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Worksheets(sheetName).ChartObjects(chartName).Activate
    ActiveChart.ChartArea.Select
	
    Select Case chartName
    Case "Chart1"
        ActiveChart.SetSourceData Source:=Worksheets(sheetName).Range("C9:" & strLastCell & "13")

    Case "Chart2"
        ActiveChart.SetSourceData Source:=Worksheets(sheetName).Range("C15:" & strLastCell & "19")

    Case "Chart3"
        ActiveChart.SetSourceData Source:=Worksheets(sheetName).Range("C21:" & strLastCell & "25")

    Case "Chart4"
        ActiveChart.SetSourceData Source:=Worksheets(sheetName).Range("C27:" & strLastCell & "38")

	Case "Chart13" '在庫
		ActiveChart.SetSourceData Source:=Worksheets(sheetName).Range("D2:" & strLastCell & "7")
    Case Else		
		
    End Select
	
    'For iRow = 1 To 3
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
    '    ActiveChart.SeriesCollection(iRow).XValues = "=" & sheetName & "!$"& strFirstCell &"$" & indexDate & ":$" & strLastCell & "$" & indexDate & ""
    'Next iRow
    
    'For iRow = 1 To 3
		'quét l?i vùng ch? d? li?u c?a bi?u dò.
	'	ActiveChart.SeriesCollection(iRow).Values = "=" & sheetName & "!$" & strFirstCell &"$" & index & ":$" & strLastCell & "$" & index
		'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    'index = index + 1
    'Next iRow
    
End Sub

Sub FillChart_GraphCFT(strFirstCell As String, strLastCell As String, indexDate As Integer, index As Integer, sheetName As String, chartName As String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Worksheets(sheetName).ChartObjects(chartName).Activate
    ActiveChart.ChartArea.Select
    
    For iRow = 1 To 2
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "=" & sheetName & "!$"& strFirstCell & "$" & indexDate & ":$" & strLastCell & "$" & indexDate & ""
    Next iRow
    
    For iRow = 1 To 2
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(iRow).Values = "=" & sheetName & "!$" & strFirstCell & "$" & index & ":$" & strLastCell & "$" & index
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    index = index + 10
    Next iRow
    
End Sub

Sub FillChart_Sheet3(strFirstCell As String, strLastCell As String, indexDate As Integer, indexActAndFree As Integer, index As Integer, sheetName As String, chartName As String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Worksheets(sheetName).ChartObjects(chartName).Activate
    ActiveChart.ChartArea.Select
    
    For iRow = 1 To 3
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "=" & sheetName & "!$" & strFirstCell & "$" & indexDate & ":$" & strLastCell & "$" & indexActAndFree & ""
    Next iRow
    
    For iRow = 1 To 3
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(iRow).Values = "=" & sheetName & "!$" & strFirstCell & "$" & index & ":$" & strLastCell & "$" & index
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    index = index + 1
    Next iRow
    
End Sub

Sub FillChart_Sheet580(strFirstCell As String, strLastCell As String, indexDate As Integer, indexActAndFree As Integer, index As Integer, sheetName As String, chartName As String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Worksheets(sheetName).ChartObjects(chartName).Activate
    ActiveChart.ChartArea.Select
    
    For iRow = 1 To 4
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "=" & sheetName & "!$" & strFirstCell & "$" & indexDate & ":$" & strLastCell & "$" & indexActAndFree & ""
    Next iRow
    
    For iRow = 1 To 4
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