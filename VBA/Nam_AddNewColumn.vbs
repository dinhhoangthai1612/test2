Sub InsertColumn()

    Dim DischargeDate As Range
    Set DischargeDate = Range("A1:AK1").Find("Bag Number")
    If DischargeDate Is Nothing Then
      MsgBox "DISCHARGE DATE column was not found."
      Exit Sub
    Else
      Columns(DischargeDate.Column).Offset(, 1).Resize(, 1).Insert
      Range("H1").Value = "MC Type"
        
    End If


End Sub
