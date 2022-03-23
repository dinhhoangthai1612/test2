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

Sub ReplaceData()
Dim myString As String
Dim ws As Worksheet
Dim var_Class As String
Dim LastRowColumnAS As Long
Dim i as Long
LastRowColumnAS = Cells(Rows.Count, 1).End(xlUp).Row
Set ws = ThisWorkbook.Sheets("T_Inventory_List")

    For i = 5 To LastRowColumnAS
    var_Class = ws.Range("AM" & CStr(i)).Value
        If (ws.Range("AN" & CStr(i)).Value = "#N/A") Then
            If (var_Class = "C") Or (var_Class = "ML") Or (var_Class = "MR") Or (var_Class = "OL") Or (var_Class = "OR") Then
                ws.Range("AN" & CStr(i)).Value = "ZIPPER_1"
            ElseIf (var_Class = "CH") Then
                ws.Range("AN" & CStr(i)).Value = "CHAIN_1"
            ElseIf (var_Class = "T") Then
                ws.Range("AN" & CStr(i)).Value = "TAPE_1"
            ElseIf (var_Class = "SP") Then
                ws.Range("AN" & CStr(i)).Value = "SLIDER PARTS_1"
            ElseIf (var_Class = "PS") Then
                 If InStr(1, Range("C" & i), "DS") <> 0 Then
                    ws.Range("AN" & CStr(i)).Value = "SLIDERD ECEC"
                 ElseIf (Not (InStr(1, myString, "DS"))) Then
                    ws.Range("AN" & CStr(i)).Value = "SLIDER_1"
                 End If
            End If
        End If
    Next
End Sub



