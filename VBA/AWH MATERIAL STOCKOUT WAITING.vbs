Sub AllProcess()

    Dim DischargeDate As Range
    Set DischargeDate = Range("A1:AK1").Find("Bag Number")
    If DischargeDate Is Nothing Then
      Exit Sub
    Else
      Columns(DischargeDate.Column).Offset(, 1).Resize(, 1).Insert
      Range("H1").Value = "MC Type"
        
    End If

Dim LastRowColumnA As Long
LastRowColumnA = Cells(Rows.Count, 1).End(xlUp).Row
Range("H2:H" & LastRowColumnA).Formula = "=IF(G2>0,""SRS"",""RM"")"

Sheets("AppCreateSheet").Range("A1:AL100000").RemoveDuplicates Columns:=6, Header:=xlYes
     



End Sub
