Sub InsertSign()
    Dim LastRow As Long
    
    Application.DisplayAlerts = False
    CheckIfSheetExists = False
    Set sht = ActiveSheet
    For Each WS In Worksheets
        If WS.Name = "DEL_Param" Then
            CheckIfSheetExists = True
            LastRow = WS.cells(sht.Rows.Count, "A").End(xlUp).Row
            For i = 1 To LastRow
                InsertPictureInRange WS.Range("C" & i).Value, WS.Range("D" & i).Value
            Next i
            WS.Delete
            Exit For
        End If
    Next WS
End Sub

Sub ChangeNamePictureEAP()

    Dim shp As Shape
    Dim shpName As String
    Dim i As Integer

    i = 0

    For Each shp In ActiveSheet.Shapes
    
        If shp.Type = msoPicture Then
        
            shpName = shp.Name
            ActiveSheet.Shapes.Range(Array(shpName)).Select
            Selection.ShapeRange.Name = "Logo " & i
            i = i + 1
            
        End If
        
    Next shp

End Sub

Sub AlignPictureEAP()
    Dim shp As Shape
    Dim shpName As String
    Dim i As Integer

    i = 0

    For Each shp In ActiveSheet.Shapes

        If shp.Type = msoPicture Then

            shpName = shp.Name

            If InStr(shpName, "Picture") > 0 Then

                ActiveSheet.Shapes.Range(Array(shpName)).Select
                Selection.ShapeRange.IncrementLeft 35.52519685
                i = i + 1

            End If

        End If

    Next shp

    cells.Replace What:="null", Replacement:="", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
    ReplaceFormat:=False

End Sub

Sub DeleteEmptyRow()

    Dim wsMain As Worksheet
    Dim i As Integer, fRow As Integer, lRow As Integer
    Dim myCell As Range
    Dim myImage As Shape
    
    For i = 1 To Workbooks.Count
        '**************************************************************************
        'ADJUST_STOCK_REQUEST_FORM
        '**************************************************************************
        If InStr(Workbooks(i).Name, "ADJUST_STOCK_REQUEST_FORM") = 1 Then
            Set wsMain = Workbooks(i).Sheets("Form")
            ' set last empty row of table - check on template
            lRow = 60
            
            ' find first empty row of table
            For j = 2 To 18
                If wsMain.cells(lRow, j).End(xlUp).Row + 1 > fRow Then
                    fRow = wsMain.cells(lRow, j).End(xlUp).Row + 1
                End If
            Next j
            
            delFlag = True
            For m = 2 To 18
                If wsMain.cells(fRow, m).Value <> "" Then
                    delFlag = False
                    Exit For
                End If
            Next m
            
            ' delete empty rows
            If delFlag Then
                wsMain.Rows(fRow & ":" & lRow).Delete
            End If
            
            
            'fRow = wsMain.cells(lRow, 4).End(xlUp).Row + 1
            ' delete empty rows
            'wsMain.Rows(fRow & ":" & lRow).Delete
            ' reset print area
            'lRow = wsMain.cells(wsMain.Rows.Count, 2).End(xlUp).Row + 1
            'wsMain.PageSetup.PrintArea = "A1:T" & lRow
            i = Workbooks.Count
        End If
        
        '**************************************************************************
        'CLAIM_REPORT
        '**************************************************************************
        If InStr(Workbooks(i).Name, "CLAIM_REPORT") = 1 Then
            Set wsMain = Workbooks(i).Sheets("Form")
            ' set last empty row of table - check on template
            lRow = 100
            ' find first empty row of table
            fRow = wsMain.cells(lRow, 2).End(xlUp).Row + 1
            ' delete empty rows
            wsMain.Rows(fRow & ":" & lRow).Delete
            ' reset print area
            lRow = wsMain.cells(wsMain.Rows.Count, 2).End(xlUp).Row + 1
            wsMain.PageSetup.PrintArea = "A1:Y" & lRow + 1
            i = Workbooks.Count
        End If
        
        '**************************************************************************
        'REQUEST_RELEASE_SHIPMENT
        '**************************************************************************
        If InStr(Workbooks(i).Name, "REQUEST_RELEASE_SHIPMENT") = 1 Then
            Set wsMain = Workbooks(i).Sheets("Form")
            ' set last empty row of table - check on template
            lRow = cells.Find(What:="REQUESTOR", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=True, SearchFormat:=False).Row - 3
            ' find first empty row of table
            'wsMain.Cells(lRow, 2).Select
            
            fRow = wsMain.cells(lRow, 1).End(xlUp).Row + 1
            
            For j = 1 To 8
                If wsMain.cells(lRow, j).End(xlUp).Row + 1 > fRow Then
                    fRow = wsMain.cells(lRow, j).End(xlUp).Row + 1
                End If
            Next j
            
            delFlag = True
            For m = 1 To 8
                If wsMain.cells(fRow, m).Value <> "" Then
                    delFlag = False
                    Exit For
                End If
            Next m
            
            
            ' delete empty rows
            If delFlag Then
                wsMain.Rows(fRow & ":" & lRow).Delete
            End If
            ' reset print area
            'lRow = wsMain.cells(wsMain.Rows.Count, 2).End(xlUp).Row + 1
            'wsMain.PageSetup.PrintArea = "A1:Y" & lRow + 1
            i = Workbooks.Count
        End If
                
        '**************************************************************************
        'SALES_PRICE_FOR_
        '**************************************************************************
        If InStr(Workbooks(i).Name, "SALES_PRICE_FOR_") = 1 Then
            Set wsMain = Workbooks(i).Sheets("Form")
            ' set last empty row of table - check on template
            lRow = cells.Find(What:="REQUESTOR", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=True, SearchFormat:=False).Row - 5
            ' find first empty row of table
            'wsMain.Cells(lRow, 2).Select
            
            fRow = wsMain.cells(lRow, 1).End(xlUp).Row + 1
            
            For j = 1 To 14
                If wsMain.cells(lRow, j).End(xlUp).Row + 1 > fRow Then
                    fRow = wsMain.cells(lRow, j).End(xlUp).Row + 1
                End If
            Next j
            
            delFlag = True
            For m = 1 To 14
                If wsMain.cells(fRow, m).Value <> "" Then
                    delFlag = False
                    Exit For
                End If
            Next m
            
            
            ' delete empty rows
            If delFlag Then
                wsMain.Rows(fRow & ":" & lRow).Delete
            End If
            i = Workbooks.Count
        End If
        
        '**************************************************************************
        'SALE_PRICE_FOR_
        '**************************************************************************
        If InStr(Workbooks(i).Name, "SALE_PRICE_FOR_") = 1 Then
            Set wsMain = Workbooks(i).Sheets("Form")
            ' set last empty row of table - check on template
            lRow = cells.Find(What:="REQUESTOR", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=True, SearchFormat:=False).Row - 3
            ' find first empty row of table
            'wsMain.Cells(lRow, 2).Select
            
            fRow = wsMain.cells(lRow, 1).End(xlUp).Row + 1
            
            For j = 1 To 14
                If wsMain.cells(lRow, j).End(xlUp).Row + 1 > fRow Then
                    fRow = wsMain.cells(lRow, j).End(xlUp).Row + 1
                End If
            Next j
            
            delFlag = True
            For m = 1 To 14
                If wsMain.cells(fRow, m).Value <> "" Then
                    delFlag = False
                    Exit For
                End If
            Next m
            
            
            ' delete empty rows
            If delFlag Then
                wsMain.Rows(fRow & ":" & lRow).Delete
            End If
            i = Workbooks.Count
        End If

        If InStr(Workbooks(i).Name, "CHANGING_FORM_(REVISE_ORDER)") = 1 Then
            lRow = Workbooks(i).Sheets("Form").Range("B28:AB78").Find(What:="*", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            Workbooks(i).Sheets("Form").Rows(lRow + 1 & ":78").Delete
            lRow = Workbooks(i).Sheets("Form").Range("MarkCell").Row
            Set myImage = Workbooks(i).Sheets("Form").Shapes("myPicture")
            Set myCell = Workbooks(i).Sheets("Form").Range("F" & lRow + 2)
            myImage.Top = myCell.Top
            myImage.Left = myCell.Left
            Workbooks(i).Save
        End If
        
        '**************************************************************************
        'INVOICE_CANCELLATION_REQUEST_FORM
        '**************************************************************************
        If InStr(Workbooks(i).Name, "INVOICE_CANCELLATION_REQUEST_FORM") = 1 Then
            Set wsMain = Workbooks(i).Sheets("Form")
            ' set last empty row of table - check on template
            lRow = 47
            
            ' find first empty row of table
            For j = 1 To 13
                If wsMain.cells(lRow, j).End(xlUp).Row + 1 > fRow Then
                    fRow = wsMain.cells(lRow, j).End(xlUp).Row + 1
                End If
            Next j
            
            delFlag = True
            For m = 1 To 13
                If wsMain.cells(fRow, m).Value <> "" Then
                    delFlag = False
                    Exit For
                End If
            Next m
            
            ' delete empty rows
            If delFlag Then
                wsMain.Rows(fRow & ":" & lRow).Delete
            End If
            
            
            'fRow = wsMain.cells(lRow, 4).End(xlUp).Row + 1
            ' delete empty rows
            'wsMain.Rows(fRow & ":" & lRow).Delete
            ' reset print area
            'lRow = wsMain.cells(wsMain.Rows.Count, 2).End(xlUp).Row + 1
            'wsMain.PageSetup.PrintArea = "A1:T" & lRow
            i = Workbooks.Count
        End If
        
        '**************************************************************************
        'INVOICE_ADJUSTMENT__REQUEST_FORM
        '**************************************************************************
        If InStr(Workbooks(i).Name, "INVOICE_ADJUSTMENT__REQUEST_FORM") = 1 Then
            Set wsMain = Workbooks(i).Sheets("Form")
            ' set last empty row of table - check on template
            lRow = 45
            
            ' find first empty row of table
            For j = 1 To 6
                If wsMain.cells(lRow, j).End(xlUp).Row + 1 > fRow Then
                    fRow = wsMain.cells(lRow, j).End(xlUp).Row + 1
                End If
            Next j
            
            delFlag = True
            For m = 1 To 6
                If wsMain.cells(fRow, m).Value <> "" Then
                    delFlag = False
                    Exit For
                End If
            Next m
            
            ' delete empty rows
            If delFlag Then
                wsMain.Rows(fRow & ":" & lRow).Delete
            End If
            
            
            'fRow = wsMain.cells(lRow, 4).End(xlUp).Row + 1
            ' delete empty rows
            'wsMain.Rows(fRow & ":" & lRow).Delete
            ' reset print area
            'lRow = wsMain.cells(wsMain.Rows.Count, 2).End(xlUp).Row + 1
            'wsMain.PageSetup.PrintArea = "A1:T" & lRow
            i = Workbooks.Count
        End If
        
        '**************************************************************************
        'CREDIT_NOTE_REQUEST_FORM
        '**************************************************************************
        If InStr(Workbooks(i).Name, "CREDIT_NOTE_REQUEST_FORM") = 1 Then
            Set wsMain = Workbooks(i).Sheets("Form")
            ' set last empty row of table - check on template
            lRow = 41
            
            ' find first empty row of table
            For j = 1 To 6
                If wsMain.cells(lRow, j).End(xlUp).Row + 1 > fRow Then
                    fRow = wsMain.cells(lRow, j).End(xlUp).Row + 1
                End If
            Next j
            
            delFlag = True
            For m = 1 To 6
                If wsMain.cells(fRow, m).Value <> "" Then
                    delFlag = False
                    Exit For
                End If
            Next m
            
            ' delete empty rows
            If delFlag Then
                wsMain.Rows(fRow & ":" & lRow).Delete
            End If
            
            
            'fRow = wsMain.cells(lRow, 4).End(xlUp).Row + 1
            ' delete empty rows
            'wsMain.Rows(fRow & ":" & lRow).Delete
            ' reset print area
            'lRow = wsMain.cells(wsMain.Rows.Count, 2).End(xlUp).Row + 1
            'wsMain.PageSetup.PrintArea = "A1:T" & lRow
            i = Workbooks.Count
        End If
        
        
        '**************************************************************************
        'DISPOSAL_REQUEST_FORM
        '**************************************************************************
        If InStr(Workbooks(i).Name, "DISPOSAL_REQUEST_FORM") = 1 Then
            Set wsMain = Workbooks(i).Sheets("Form")
            ' set last empty row of table - check on template
            lRow = 55
            
            ' find first empty row of table
            For j = 1 To 12
                If wsMain.cells(lRow, j).End(xlUp).Row + 1 > fRow Then
                    fRow = wsMain.cells(lRow, j).End(xlUp).Row + 1
                End If
            Next j
            
            delFlag = True
            For m = 1 To 12
                If wsMain.cells(fRow, m).Value <> "" Then
                    delFlag = False
                    Exit For
                End If
            Next m
            
            ' delete empty rows
            If delFlag Then
                wsMain.Rows(fRow & ":" & lRow).Delete
            End If
            i = Workbooks.Count
        End If
        
    Next i

End Sub



Sub AdjustForm()
    Application.DisplayAlerts = False
    CheckIfSheetExists = False
    CheckIfConfirmExists = False
    Dim sheetFrom As String, rangeFrom As String
    For Each WS In Worksheets
        If WS.Name = "CopyData" Then
            CheckIfSheetExists = True
        End If
        If WS.Name = "CONFIRM" Then
            CheckIfConfirmExists = True
        End If
    Next WS
    
    'copy format
    If CheckIfSheetExists Then
        LastRow = Sheets("CopyData").cells(Sheets("CopyData").Rows.Count, "A").End(xlUp).Row
        For i = 2 To LastRow
            sheetFrom = Sheets("CopyData").Range("A" & i).Text
            rangeFrom = Sheets("CopyData").Range("B" & i).Text
            If sheetFrom <> "" And rangeFrom <> "" Then
                CopyRange sheetFrom, rangeFrom
            End If
        Next i
        
        
        Sheets("CopyData").Delete
        'MsgBox "Delete CopyData"
        
    End If
    
    If CheckIfConfirmExists Then
    
        For i = 1 To 5
            If Sheets("CONFIRM").Range("A" & i).Value <> "" Then
            
                Sheets("CONFIRM").Select
                arrTarget = Split(Sheets("CONFIRM").Range("A" & i).Value, ":")
                strFind = Sheets("CONFIRM").Range(arrTarget(0)).Value
                colFind = Sheets("CONFIRM").Range(arrTarget(0)).Column
                
                Range(Sheets("CONFIRM").Range("A" & i).Value).Select
                Selection.Copy
                
                
                Sheets(1).Select
                Sheets(1).Columns(colFind).Select
                On Error GoTo Oops:
                targetCell = Selection.Find(What:=strFind, After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Address
COPYTO:
                If targetCell <> "ERR" Then
                    Range(targetCell).Select
                    ActiveSheet.Paste
                End If

            End If
        Next i
        
        Sheets("CONFIRM").Delete
        'MsgBox "Delete CONFIRM"
        
    End If

Oops:
    'handle error here
    If i < 6 And CheckIfConfirmExists Then
        targetCell = "ERR"
        Resume COPYTO 'risk of endless loop if the new URL is also bad
    End If
End Sub

Sub CopyRange(sheetFrom As String, rangeFrom As String)
    Sheets(sheetFrom).Select
    Sheets(sheetFrom).Range(rangeFrom).Copy
    Sheets(1).Select
    rangeTo = Split(rangeFrom, ":")(0)
    Sheets(1).Range(rangeTo).Select
    ActiveSheet.Paste
    Sheets(1).Range("A1").Select
End Sub

Sub InsertPictureInRange(PictureFileName As String, cells As String)
'Worksheets(sNum).Select
Dim TargetCells As Range
Dim p As Object, t As Double, l As Double, w As Double, h As Double
Dim cellF, cellT As Range, colMerge As Double



Set TargetCells = Range(cells)

    For i = 1 To 2
        Set cellF = TargetCells.Offset(0 - i, 0)
        If cellF.Value <> "" Then
            i = 10
        End If
    Next i

    For i = 1 To 5
        Set cellT = TargetCells.Offset(i, 0)
        If cellT.Value <> "" Then
            i = 10
        End If
    Next i



'MsgBox cellF.Value & "___" & cellT.Value
'MsgBox cellF.MergeArea.Columns.Count & "___" & cellT.MergeArea.Columns.Count

colMerge = cellF.MergeArea.Columns.Count
If cellT.MergeArea.Columns.Count > cellF.MergeArea.Columns.Count Then
    colMerge = cellT.MergeArea.Columns.Count
End If

'MsgBox cellT.MergeArea.cells(0, colMerge).Address
'MsgBox Range("AJ5:AK7").Height & "___" & Range("AJ5:AK7").Width

Set TargetCells = Range(cells & ":" & cellT.MergeArea.cells(0, colMerge).Address)
With TargetCells
    t = .Top
    l = .Left
    w = .Offset(0, .Columns.Count).Left - .Left
    h = .Offset(.Rows.Count, 0).Top - .Top
End With

Set p = TargetCells.Worksheet.Pictures.Insert(PictureFileName)

'MsgBox TargetCells.Worksheet.Name

With p
    .Width = w - 5
    If .Height > h Then
        .Height = h
    End If
    .Top = t
    .Left = 6 + (w - p.Width) / 2
    .Name = PictureFileName
End With
    
Set Pic = TargetCells.Worksheet.Shapes.AddPicture(PictureFileName, _
linktofile:=msoFalse, savewithdocument:=msoCTrue, Left:=0, Top:=0, _
Width:=p.Width * 10, Height:=p.Height * 10)

With Pic
    .Height = p.Height
    .Width = p.Width - 5
    .Top = t
    .Left = (l + (w - p.Width) / 2) + 2.5
End With

p.Delete

End Sub




