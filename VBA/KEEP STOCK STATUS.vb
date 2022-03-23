Sub InputDataColumnE()
    Dim myString As String
    Dim ws As Worksheet
    Dim var_Class As String
    Dim var_Class2 As String
    Dim var_Class3 As String
    Dim LastRowColumnAS As Long
    Dim i As Long

    LastRowColumnAS = Cells(Rows.Count, 1).End(xlUp).Row
    Set ws = ThisWorkbook.Sheets("Stock")

    For i = 3 To LastRowColumnAS
        var_Class = ws.Range("D" & CStr(i)).Value
        var_Class2 = ws.Range("B" & CStr(i)).Value
        var_Class3 = ws.Range("E" & CStr(i)).Value
        var_Class = Trim(var_Class)
        var_Class2 = Trim(var_Class2)
        var_Class3 = Trim(var_Class3)
            If (var_Class = "PS") Then
                ws.Range("E" & CStr(i)).Value = "SLIDER"
            ElseIf (var_Class = "CH") And InStr(var_Class2, "CNT") Or InStr(var_Class2, "CFT") Then
                ws.Range("E" & CStr(i)).Value = "CHAIN AQUA GUARD"
            ElseIf (var_Class = "CH") And (var_Class3 = "") Then
                ws.Range("E" & CStr(i)).Value = "CHAIN"
            ElseIf (var_Class = "C") Or (var_Class = "OL") Or (var_Class = "OR") Or (var_Class = "ML") Or (var_Class = "MR") Then
                ws.Range("E" & CStr(i)).Value = "CUT ZIPPER"
            ElseIf (var_Class3 = "") Then
                ws.Range("E" & CStr(i)).Value = "OTHERS"
            End If
    Next
End Sub

Sub InputDataColumnF()
    Dim myString As String
    Dim ws As Worksheet
    Dim var_Class As String
    Dim LastRowColumnAS As Long
    Dim i As Long

    LastRowColumnAS = Cells(Rows.Count, 1).End(xlUp).Row
    Set ws = ThisWorkbook.Sheets("Stock")

    For i = 3 To LastRowColumnAS
        var_Class = ws.Range("F" & CStr(i)).Value
        var_Class = Trim(var_Class)
            If (var_Class = "1") Then
                ws.Range("F" & CStr(i)).Value = "PRO"
            ElseIf (var_Class = "") Then
                ws.Range("F" & CStr(i)).Value = "PUR"
            End If
    Next
End Sub

Sub CalculateColumnM()
    Dim myString As String
    Dim ws As Worksheet
    Dim var_Class As String
    Dim var_Class2 As String
    Dim var_Class3 As String
    Dim LastRowColumnAS As Long
    Dim i As Long

    LastRowColumnAS = Cells(Rows.Count, 1).End(xlUp).Row
    Set ws = ThisWorkbook.Sheets("Stock")

    For i = 3 To LastRowColumnAS
        var_Class = ws.Range("H" & CStr(i)).Value
        var_Class2 = ws.Range("W" & CStr(i)).Value
        var_Class3 = ws.Range("X" & CStr(i)).Value
        var_Class = Trim(var_Class)
        var_Class2 = Trim(var_Class2)
        var_Class3 = Trim(var_Class3)
            If (var_Class = "C") Then
                ws.Range("M" & CStr(i)).Value = "=+((G" & CStr(i) & "/100)*W" & CStr(i) & "+X" & CStr(i) & ")*K" & CStr(i)
            ElseIf (var_Class = "I") Then
                ws.Range("M" & CStr(i)).Value = "=+((G" & CStr(i) & "*2.54/100)*W" & CStr(i) & "+X" & CStr(i) & ")*K" & CStr(i)
            ElseIf (var_Class2 = "0") Then
                ws.Range("M" & CStr(i)).Value = "=+K" & CStr(i) & "*X" & CStr(i)
            ElseIf (var_Class3 = "0") Then
                ws.Range("M" & CStr(i)).Value = "=+K" & CStr(i) & "*W" & CStr(i)
            End If
    Next
End Sub

Function InsertColumn(getColumn as String, sheetName as String)
    Application.DisplayAlerts = False
    Worksheets(sheetName).Activate
    Columns(getColumn & ":" & getColumn).Select
    Selection.Insert Shift:=xlToRight
End Function

Sub UpdateChart1(strFirstCell As String, strLastCell As String, indexDate As Integer, index As Integer, sheetName As String, chartName As String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Dim countRow As Integer
    Worksheets(sheetName).ChartObjects(chartName).Activate
    ActiveChart.ChartArea.Select
    
    countRow = 6
    For iRow = 1 To countRow
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "=" & sheetName & "!$"& strFirstCell &"$" & indexDate & ":$" & strLastCell & "$" & indexDate & ""
    Next iRow
    
    countRow = 6
    For iRow = 1 To countRow
        If (countRow = 1) Then
            index = 41
        End If
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(iRow).Values = "=" & sheetName & "!$" & strFirstCell &"$" & index & ":$" & strLastCell & "$" & index
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    countRow = countRow - 1
    index = index + 1
    Next iRow
End Sub

Sub UpdateChart2(strFirstCell As String, strLastCell As String, indexDate As Integer, index As Integer, sheetName As String, chartName As String)
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

Sub UpdateChart3(strFirstCell As String, strLastCell As String, indexDate As Integer, index As Integer, sheetName As String, chartName As String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Worksheets(sheetName).ChartObjects(chartName).Activate
    ActiveChart.ChartArea.Select
    
    For iRow = 1 To 7
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "=" & sheetName & "!$"& strFirstCell &"$" & indexDate & ":$" & strLastCell & "$" & indexDate & ""
    Next iRow
    
    For iRow = 1 To 7
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(iRow).Values = "=" & sheetName & "!$" & strFirstCell &"$" & index & ":$" & strLastCell & "$" & index
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
    index = index + 1
    Next iRow
End Sub

Sub UpdateChart4(strFirstCell As String, strLastCell As String, indexDate As Integer, index As Integer, sheetName As String, chartName As String)
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

Sub UpdateChart5(strFirstCell As String, strLastCell As String, indexDate As Integer, index As Integer, sheetName As String, chartName As String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Worksheets(sheetName).ChartObjects(chartName).Activate
    ActiveChart.ChartArea.Select
    
    For iRow = 1 To 2
        ' quét l?i vùng d? li?u ngày tháng nam c?a t?ng c?t bi?u d?.
        ActiveChart.SeriesCollection(iRow).XValues = "=" & sheetName & "!$"& strFirstCell &"$" & indexDate & ":$" & strLastCell & "$" & indexDate & ""
    Next iRow
    
    For iRow = 1 To 2
    'quét l?i vùng ch? d? li?u c?a bi?u dò.
    ActiveChart.SeriesCollection(iRow).Values = "=" & sheetName & "!$" & strFirstCell &"$" & index & ":$" & strLastCell & "$" & index
    'ActiveChart.SeriesCollection(index).Name = "='Report'!$D$" & iRow
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

Sub ChangeTitleTopSalesChart(monthChart as String, yearChart as String)
    Sheets("TOP SALES").ChartObjects("Chart 1").Chart.ChartTitle.Text = "Keep Stock by Salesperson End " & monthChart & "'" & yearChart
End Sub

Sub ChangeTitleOfficeDivChart(monthChart as String, yearChart as String)
    Sheets("BY OFFICE-DIV").ChartObjects("Chart 1").Chart.ChartTitle.Text = "Keep Stock by Office/Div End " & monthChart & "'" & yearChart
End Sub

Sub UpdateChart_PRD_PUR(sheetName As String, chartName As String, strCellValue As String)
    Application.DisplayAlerts = False
    'sheetName As String, chartName As String, strCellValue As String
    Dim strFirstCell As String
    strFirstCell = "D"
    
    Worksheets(sheetName).ChartObjects(chartName).Activate
    ActiveChart.ChartArea.Select
    
    ActiveChart.SeriesCollection(1).XValues = "='" & sheetName & "'!$D$25:$" & strCellValue & "$25"
    ActiveChart.SeriesCollection(2).XValues = "='" & sheetName & "'!$D$25:$" & strCellValue & "$25"

    ActiveChart.SeriesCollection(1).Values = "='" & sheetName & "'!$D$26:$" & strCellValue & "$26"
    ActiveChart.SeriesCollection(2).Values = "='" & sheetName & "'!$D$27:$" & strCellValue & "$27"
End Sub


