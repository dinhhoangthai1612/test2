Sub WorkbookRefreshQuery()
    Application.DisplayAlerts = False
    ActiveWorkbook.RefreshAll
End Sub

Sub RemoveFormular()
    Dim ws As Worksheet
    Dim LastRowColumnAS As Long
    LastRowColumnAS = Cells(Rows.Count, 1).End(xlUp).Row

    Set ws = ThisWorkbook.Sheets("T_Inventory_List")

    With ws.Range("A5:AS" & LastRowColumnAS)
        .Copy
        .PasteSpecial Paste:=xlPasteValues, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End With
    

End Sub


Sub Filter1()
Dim myString As String
Dim ws As Worksheet
Dim var_Class As String
Dim LastRowColumnAS As Long
Dim i As Long

LastRowColumnAS = Cells(Rows.Count, 1).End(xlUp).Row
Set ws = ThisWorkbook.Sheets("T_Inventory_List")

    For i = 5 To LastRowColumnAS
    var_Class = ws.Range("AM" & CStr(i)).Value
    var_Class = Trim(var_Class)
            If (var_Class = "C") Or (var_Class = "ML") Or (var_Class = "MR") Or (var_Class = "OL") Or (var_Class = "OR") Then
                ws.Range("AN" & CStr(i)).Value = "ZIPPER"
            ElseIf (var_Class = "CH") Then
                ws.Range("AN" & CStr(i)).Value = "CHAIN"
            ElseIf (var_Class = "T") Then
                ws.Range("AN" & CStr(i)).Value = "TAPE"
            ElseIf (var_Class = "SP") Then
                ws.Range("AN" & CStr(i)).Value = "SLIDER PARTS"
            ElseIf (var_Class = "PS") Then
                 If InStr(1, Range("C" & i), " DS") > 0 Then
                    ws.Range("AN" & CStr(i)).Value = "SLIDER DS*"
                 Else
                    ws.Range("AN" & CStr(i)).Value = "SLIDER"
                 End If
            ElseIf (var_Class = "EL") Then
                ws.Range("AN" & CStr(i)).Value = "CF ELEMENT"
            ElseIf (var_Class = "FD") Then
                ws.Range("AN" & CStr(i)).Value = "CHEMICAL"
            ElseIf (var_Class = "FF") or var_Class = "FR" or var_Class = "FY" Then
                ws.Range("AN" & CStr(i)).Value = "FILM"
            ElseIf (var_Class = "GN") Then
                ws.Range("AN" & CStr(i)).Value = "DYESTUFFS"
            ElseIf (var_Class = "MF") Then
                ws.Range("AN" & CStr(i)).Value = "MONOFILA"
            ElseIf (var_Class = "PO") OR var_Class = "PB" OR var_Class = "PT" OR var_Class = "WB" Then
                ws.Range("AN" & CStr(i)).Value = "ZIPPER PARTS"
            ElseIf (var_Class = "Q") Then
                ws.Range("AN" & CStr(i)).Value = "COSMOLON"
            ElseIf (var_Class = "ST") OR var_Class = "TY" Then
                ws.Range("AN" & CStr(i)).Value = "YARN"   
            ElseIf (var_Class = "TC") Then
                ws.Range("AN" & CStr(i)).Value = "CORD"     
            ElseIf (var_Class = "WE") Then
                ws.Range("AN" & CStr(i)).Value = "ELEMENT WIRE"       
            ElseIf (var_Class = "WS") Then
                ws.Range("AN" & CStr(i)).Value = "SLIDER WIRE"                
            Else
                ws.Range("AN" & CStr(i)).Value = "BLANK"
            End If
    Next
End Sub



Sub ReplaceData2()
Dim ws As Worksheet
Dim var_Class As String
Dim var_Class2 As String
Dim LastRowColumnAS As Long
Dim i As Long

LastRowColumnAS = Cells(Rows.Count, 1).End(xlUp).Row

Set ws = ThisWorkbook.Sheets("T_Inventory_List")

    For i = 5 To LastRowColumnAS
        var_Class = CStr(Range("N" & CStr(i)).Value)
        If CStr(var_Class) = "BLANK" or CStr(var_Class) = "#N/A" Then
            Range("N" & CStr(i)) = ">=19"
        End If
    Next 

    For i = 5 To LastRowColumnAS
        var_Class = CStr(Range("Y" & CStr(i)).Value)
        If CStr(var_Class) = "0" Then
        Range("AA" & CStr(i)) = Cells(i, 10) * Cells(i, 26)
        End If
    Next 

    For i = 5 To LastRowColumnAS
        var_Class = CStr(Range("Z" & CStr(i)).Value)
        If CStr(var_Class) = "0" Then
        Range("AA" & CStr(i)) = Cells(i, 10) * Cells(i, 25)
        End If
    Next 

    For i = 5 To LastRowColumnAS
        var_Class = CStr(Range("H" & CStr(i)).Value)
        If CStr(var_Class) = "C" Then
        Range("AA" & CStr(i)) = (Cells(i,7) / 100 * cells(i,25) + cells(i, 26)) * cells(i,10)
        End If
    Next 

    For i = 5 To LastRowColumnAS
        var_Class = CStr(Range("H" & CStr(i)).Value)
        If CStr(var_Class) = "I" Then
        Range("AA" & CStr(i)) = (Cells(i,7) * 2.54 / 100 * cells(i,25) + cells(i, 26)) * cells(i,10)
        End If
    Next 

    For i = 5 To LastRowColumnAS
    var_Class = Range("AM" & (i)).Value
        If Trim(var_Class) = "PS" Then
            'Debug.Print "CLSCDC: " & var_Class
            var_Class2 = Range("AP" & (i)).Value
            If Trim(var_Class2) = "ET" Or Trim(var_Class2) Like "Z*" Then
                'Debug.Print "COLUMN AP: " & var_Class2
                Range("L" & (i)) = "0"
            End If
        End If
    Next
End Sub


Sub ConvertToText()

Dim vData As Variant
vData = Columns(1)
Columns(1).NumberFormat = "@"
Columns(1) = vData

vData = Columns(2)
Columns(2).NumberFormat = "@"
Columns(2) = vData
End Sub

Sub RemoveDuplicates()
Sheets("class").Range("A1:F1000000").RemoveDuplicates Columns:=2, Header:=xlYes
End Sub

Function ChangeAllCharts(nowColum As String, nextColumn As String)

    Dim sht As Worksheet

    Dim cht As ChartObject

    Dim ser As Series
    
    

    For Each sht In ActiveWorkbook.Worksheets

        For Each cht In sht.ChartObjects

            For Each ser In cht.Chart.SeriesCollection

                ser.Formula = Replace(ser.Formula, nowColum, nextColumn)

            Next ser

        Next cht

    Next sht

End Function

