Sub WorkbookRefreshQuery()
    Application.DisplayAlerts = False
    ActiveWorkbook.RefreshAll
End Sub

Function InsertColumn(getColumn as String, sheetName as String)
    Application.DisplayAlerts = False
    Worksheets(sheetName).Activate
    Columns(getColumn & ":" & getColumn).Select
    Selection.Insert Shift:=xlToRight
End Function

Function HideColumn(getColumn as String)
    Columns(getColumn).Hidden = True
    Application.Wait (Now + TimeValue("0:00:01"))
End Function

Sub FillChart_graph(strCellValue as String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Dim index As Integer

    Worksheets("graph").ChartObjects("Chart 1").Activate
    ActiveChart.ChartArea.Select
    
    For iRow = 1 To 16
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "='graph'!$AM$10:$" & strCellValue & "$11"
    Next iRow

    index = 12
    For iRow = 1 To 16
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(iRow).Values = "='graph'!$AM$" & index & ":$" & strCellValue & "$" & index
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    index = index + 1
    Next iRow

    Worksheets("graph").ChartObjects("Chart 2").Activate
    ActiveChart.ChartArea.Select

    '''''''''''''''''''''''''

    For iRow = 1 To 16
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "='graph'!$AZ$41:$" & strCellValue & "$42"
    Next iRow

    index = 43
    For iRow = 1 To 16
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(iRow).Values = "='graph'!$AZ$" & index & ":$" & strCellValue & "$" & index
    index = index + 1
    Next iRow

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


Sub FillChart_ReportSheet(strCellValue as String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Dim index As Integer

    Worksheets("Report").ChartObjects("Chart 3").Activate
    ActiveChart.ChartArea.Select
    
    For iRow = 1 To 5
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "='Report'!$D$2:$" & strCellValue & "$2"
    Next iRow

    index = 3
    For iRow = 1 To 5
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(iRow).Values = "='Report'!$D$" & index & ":$" & strCellValue & "$" & index
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    index = index + 1
    Next iRow

    Worksheets("Report").ChartObjects("Chart 4").Activate
    ActiveChart.ChartArea.Select

    '''''''''''''''''''''''''

    For iRow = 1 To 5
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "='Report'!$D$12:$" & strCellValue & "$12"
    Next iRow

    index = 13
    For iRow = 1 To 5
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(iRow).Values = "='Report'!$D$" & index & ":$" & strCellValue & "$" & index
    index = index + 1
    Next iRow

End Sub


Sub ChangeTitle_DValue(datetimeChart as String)
Sheets("graph").ChartObjects("Chart 1").Chart.ChartTitle.Text = "D_VALUE SITUATION " & datetimeChart & "(BY Q'TY)"
Sheets("graph").ChartObjects("Chart 2").Chart.ChartTitle.Text = "D_VALUE SITUATION " & datetimeChart & "(BY PR)"
End Sub

Sub ChangeTitle_Remake(datetimeChart as String)
Sheets("report").ChartObjects("Chart 3").Chart.ChartTitle.Text = "REMAKE SITUATION  " & datetimeChart & "(BY Q'TY)"
Sheets("report").ChartObjects("Chart 4").Chart.ChartTitle.Text = "REMAKE SITUATION  " & datetimeChart & "(BY PR)"
End Sub

Sub DeleteValue_ColumnN()
Dim myString As String
Dim ws As Worksheet
Dim var_Class As String
Dim LastRowColumnAS As Long
Dim i As Long

LastRowColumnAS = Cells(Rows.Count, 1).End(xlUp).Row

Set ws = ThisWorkbook.Sheets("Source")

    For i = 2 To LastRowColumnAS
    
    var_Class = ws.Range("D" & CStr(i)).Value
    
    var_Class = Trim(var_Class)
    
            If (var_Class = "92") Then
            
                ws.Range("N" & CStr(i)).Value = ""
                
            End If
    Next
    
End Sub

Sub FitColumn(get_Column as String)
Worksheets("Report").Columns(get_Column & ":" & get_Column).AutoFit
End Sub



