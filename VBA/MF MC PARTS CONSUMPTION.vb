Sub ConvertToText()
 Dim vData As Variant

 vData = Columns(15)
 Columns(15).NumberFormat = "0"
 Columns(15) = vData

End Sub

Sub AutoFillValue(src As String, dest as String)
Set ws = ThisWorkbook.Sheets("Assembly POP")
Dim out As Range
Dim source as Range

Set out = ws.Range(dest)
Set source = ws.Range(src)
source.AutoFill Destination:=out, Type:=xlFillCopy
End Sub

Sub RemoveFormular()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets("Record Sheet")

    With ws.Range("C17:Z24")
        .Copy
        .PasteSpecial Paste:=xlPasteValues, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End With

    With ws.Range("C29:Z36")
        .Copy
        .PasteSpecial Paste:=xlPasteValues, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End With

    With ws.Range("AD17:AO20")
        .Copy
        .PasteSpecial Paste:=xlPasteValues, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End With

    With ws.Range("AD29:AO32")
        .Copy
        .PasteSpecial Paste:=xlPasteValues, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End With

End Sub