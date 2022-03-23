Sub FillChart_graph()
    Dim WS_Count As Integer
    Dim i As Integer
    Dim cht As Chart
    Dim sFormula As String
    Dim rFirst As Range, rLast As Range
    Dim LastColumnHide, LastColumnLetter As String
    Dim LastCol As Long
    
    ' Set WS_Count equal to the number of worksheets in the active
    ' workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    ' Begin the loop.
    For i = 1 To WS_Count
        If InStr(ActiveWorkbook.Worksheets(i).Name, "5CI)_Rev") > 0 Then
            Worksheets(i).Select
            LastColumnLetter = Split(Worksheets(i).cells(9, Worksheets(i).Columns.Count).End(xlToLeft).Address, "$")(1)
            
            LastCol = Worksheets(i).cells(9, Worksheets(i).Columns.Count).End(xlToLeft).Column
            LastColumnHide = Split(cells(1, LastCol - (24 * 2) + 1).Address, "$")(1)
            
            Worksheets(i).ChartObjects("3CF").Activate
            ActiveChart.SetSourceData Source:=Worksheets(i).Range("E2:E15," & LastColumnHide & "2:" & LastColumnLetter & "15")
            
            Worksheets(i).ChartObjects("5CI").Activate
            ActiveChart.SetSourceData Source:=Worksheets(i).Range("E2:E15," & LastColumnHide & "2:" & LastColumnLetter & "15")
            
            Worksheets(i).ChartObjects("3Y").Activate
            ActiveChart.SetSourceData Source:=Worksheets(i).Range("E2:E15," & LastColumnHide & "2:" & LastColumnLetter & "15")
            
            LastColumnHide = Split(cells(1, LastCol - (24 * 2)).Address, "$")(1)
            Worksheets(i).Columns("F:" & LastColumnHide).EntireColumn.Hidden = True
        
        End If
        
        If InStr(ActiveWorkbook.Worksheets(i).Name, "5CI) _Rev") > 0 Then
            '='???? (3CF?5CI) _Rev'!$C$11:$D$17,'???? (3CF?5CI) _Rev'!$AP$11:$BI$17
            Worksheets(i).Select
            
            LastColumnLetter = Split(Worksheets(i).cells(12, Worksheets(i).Columns.Count).End(xlToLeft).Address, "$")(1)
            LastCol = Worksheets(i).cells(11, Worksheets(i).Columns.Count).End(xlToLeft).Column
            
            LastColumnHide = Split(cells(1, LastCol - 20 + 1).Address, "$")(1)
            
            Worksheets(i).ChartObjects("NAT TAPE").Activate
            ActiveChart.SetSourceData Source:=Worksheets(i).Range("C11:D17," & LastColumnHide & "11:" & LastColumnLetter & "17")
            
            Worksheets(i).ChartObjects("NAT CHAIN").Activate
            ActiveChart.SetSourceData Source:=Worksheets(i).Range("C19:D23," & LastColumnHide & "19:" & LastColumnLetter & "23")
            
            Worksheets(i).ChartObjects("SET TAPE").Activate
            ActiveChart.SetSourceData Source:=Worksheets(i).Range("C35:D37," & LastColumnHide & "35:" & LastColumnLetter & "37")
            
            Worksheets(i).ChartObjects("SET CHAIN").Activate
            ActiveChart.SetSourceData Source:=Worksheets(i).Range("C39:D43," & LastColumnHide & "39:" & LastColumnLetter & "43")
            
            LastColumnHide = Split(cells(1, LastCol - 20).Address, "$")(1)
            Worksheets(i).Columns("E:" & LastColumnHide).EntireColumn.Hidden = True
        End If
        
    Next i
    
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


Sub FillChart()
    Dim WS_Count As Integer
    Dim sheetName As Integer
    Dim index As Integer
    Dim LastColumnHide, LastColumnLetter As String
    Dim LastCol As Long
    Dim iRow As Integer
    Application.DisplayAlerts = False
    
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    For sheetName = 1 To WS_Count
        If InStr(ActiveWorkbook.Worksheets(sheetName).Name, "(Tape)") > 0 Then
            Worksheets(sheetName).Select
            LastColumnLetter = Split(Worksheets(sheetName).Cells(3, Worksheets(sheetName).Columns.Count).End(xlToLeft).Address, "$")(1)
            
            LastCol = Worksheets(sheetName).Cells(3, Worksheets(sheetName).Columns.Count).End(xlToLeft).Column
            LastColumnHide = Split(Cells(1, LastCol - 20 + 1).Address, "$")(1)
            
            'Chart 1
            Worksheets(sheetName).ChartObjects("Chart 1").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "3:" & LastColumnLetter & "3")
            index = 32
            For iRow = 1 To 2
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'Chart 2
            Worksheets(sheetName).ChartObjects("Chart 2").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "3:" & LastColumnLetter & "3")
            index = 64
            For iRow = 1 To 2
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'Chart 3
            Worksheets(sheetName).ChartObjects("Chart 3").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "3:" & LastColumnLetter & "3")
            index = 4
            For iRow = 1 To 6
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'Chart 4
            Worksheets(sheetName).ChartObjects("Chart 4").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "3:" & LastColumnLetter & "3")
            index = 10
            For iRow = 1 To 10
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'Chart 5
            Worksheets(sheetName).ChartObjects("Chart 5").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "3:" & LastColumnLetter & "3")
            index = 20
            For iRow = 1 To 12
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'Chart 6
            Worksheets(sheetName).ChartObjects("Chart 6").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "35:" & LastColumnLetter & "35")
            index = 36
            For iRow = 1 To 6
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'Chart 7
            Worksheets(sheetName).ChartObjects("Chart 7").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "35:" & LastColumnLetter & "35")
            index = 42
            For iRow = 1 To 10
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'Chart 8
            Worksheets(sheetName).ChartObjects("Chart 8").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "35:" & LastColumnLetter & "35")
            index = 52
            For iRow = 1 To 12
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'An cot du lieu cu
            LastColumnHide = Split(Cells(1, LastCol - 20).Address, "$")(1)
            Worksheets(sheetName).Columns("E:" & LastColumnHide).EntireColumn.Hidden = True
        End If

        If InStr(ActiveWorkbook.Worksheets(sheetName).Name, "(CFCh)") > 0 Then
            Worksheets(sheetName).Select
            'Lay vi tri cot co du lieu cuoi cung
            LastColumnLetter = Split(Worksheets(sheetName).Cells(3, Worksheets(sheetName).Columns.Count).End(xlToLeft).Address, "$")(1)
            
            LastCol = Worksheets(sheetName).Cells(3, Worksheets(sheetName).Columns.Count).End(xlToLeft).Column
            LastColumnHide = Split(Cells(1, LastCol - 20 + 1).Address, "$")(1)
            
            'Chart 9
            Worksheets(sheetName).ChartObjects("Chart 9").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "3:" & LastColumnLetter & "3")
            index = 24
            For iRow = 1 To 2
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'Chart 10
            Worksheets(sheetName).ChartObjects("Chart 10").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "3:" & LastColumnLetter & "3")
            index = 48
            For iRow = 1 To 2
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'Chart 11
            Worksheets(sheetName).ChartObjects("Chart 11").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "3:" & LastColumnLetter & "3")
            index = 72
            For iRow = 1 To 2
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'Chart 12
            Worksheets(sheetName).ChartObjects("Chart 12").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "3:" & LastColumnLetter & "3")
            index = 4
            For iRow = 1 To 9
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'Chart 13
            Worksheets(sheetName).ChartObjects("Chart 13").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "3:" & LastColumnLetter & "3")
            index = 13
            For iRow = 1 To 11
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'Chart 14
            Worksheets(sheetName).ChartObjects("Chart 14").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "27:" & LastColumnLetter & "27")
            index = 28
            For iRow = 1 To 9
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'Chart 15
            Worksheets(sheetName).ChartObjects("Chart 15").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "27:" & LastColumnLetter & "27")
            index = 37
            For iRow = 1 To 11
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'Chart 16
            Worksheets(sheetName).ChartObjects("Chart 16").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "51:" & LastColumnLetter & "51")
            index = 52
            For iRow = 1 To 9
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            'Chart 17
            Worksheets(sheetName).ChartObjects("Chart 17").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "51:" & LastColumnLetter & "51")
            index = 61
            For iRow = 1 To 11
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            LastColumnHide = Split(Cells(1, LastCol - 20).Address, "$")(1)
            Worksheets(sheetName).Columns("E:" & LastColumnHide).EntireColumn.Hidden = True
        End If
        
        
        If InStr(ActiveWorkbook.Worksheets(sheetName).Name, "(_Tape)") > 0 Then
            Worksheets(sheetName).Select
            LastColumnLetter = Split(Worksheets(sheetName).Cells(9, Worksheets(sheetName).Columns.Count).End(xlToLeft).Address, "$")(1)
            
            LastCol = Worksheets(sheetName).Cells(9, Worksheets(sheetName).Columns.Count).End(xlToLeft).Column
            LastColumnHide = Split(Cells(1, LastCol - (24 * 2) + 1).Address, "$")(1)
            
            Worksheets(sheetName).ChartObjects("Chart 18").Activate
            ActiveChart.SetSourceData Source:=Worksheets(sheetName).Range("E3:E31," & LastColumnHide & "3:" & LastColumnLetter & "31")
            
            Worksheets(sheetName).ChartObjects("Chart 19").Activate
            ActiveChart.SetSourceData Source:=Worksheets(sheetName).Range("E3:E31," & LastColumnHide & "3:" & LastColumnLetter & "31")
            
            Worksheets(sheetName).ChartObjects("Chart 20").Activate
            ActiveChart.SetSourceData Source:=Worksheets(sheetName).Range("E3:E31," & LastColumnHide & "3:" & LastColumnLetter & "31")
            
            Worksheets(sheetName).ChartObjects("Chart 21").Activate
            'quet lai vung chon ngay cua bieu do
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "2:" & LastColumnLetter & "2")
            index = 32
            For iRow = 1 To 28
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            'ActiveChart.SetSourceData Source:=Worksheets(sheetName).Range("E30:E55," & LastColumnHide & "30:" & LastColumnLetter & "55")
            
            LastColumnHide = Split(Cells(1, LastCol - (24 * 2)).Address, "$")(1)
            Worksheets(sheetName).Columns("F:" & LastColumnHide).EntireColumn.Hidden = True
        End If
        
        If InStr(ActiveWorkbook.Worksheets(sheetName).Name, "(_CFCh)") > 0 Then
            Worksheets(sheetName).Select
            LastColumnLetter = Split(Worksheets(sheetName).Cells(9, Worksheets(sheetName).Columns.Count).End(xlToLeft).Address, "$")(1)
            
            LastCol = Worksheets(sheetName).Cells(9, Worksheets(sheetName).Columns.Count).End(xlToLeft).Column
            LastColumnHide = Split(Cells(1, LastCol - (24 * 2) + 1).Address, "$")(1)
            
            Worksheets(sheetName).ChartObjects("Chart 22").Activate
            ActiveChart.SetSourceData Source:=Worksheets(sheetName).Range("E2:E12," & LastColumnHide & "2:" & LastColumnLetter & "12")
            
            Worksheets(sheetName).ChartObjects("Chart 23").Activate
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "2:" & LastColumnLetter & "3")
            index = 13
            For iRow = 1 To 11
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
            
            Worksheets(sheetName).ChartObjects("Chart 24").Activate
            ActiveChart.SeriesCollection(1).XValues = Worksheets(sheetName).Range(LastColumnHide & "2:" & LastColumnLetter & "2")
            index = 26
            For iRow = 1 To 20
            'quet lai vung chon du lieu cua bieu do
                ActiveChart.SeriesCollection(iRow).Values = Worksheets(sheetName).Range(LastColumnHide & index & ":" & LastColumnLetter & index)
                index = index + 1
            Next iRow
                     
            LastColumnHide = Split(Cells(1, LastCol - (24 * 2)).Address, "$")(1)
            Worksheets(sheetName).Columns("F:" & LastColumnHide).EntireColumn.Hidden = True
        End If
    Next sheetName

End Sub


