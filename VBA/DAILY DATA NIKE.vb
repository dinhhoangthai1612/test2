Sub UnHideColumn()
    Dim lastColumn As Integer
    lastColumn = 40
    For i = 1 To lastColumn
        If Columns(i).EntireColumn.Hidden = True Then
            Columns(i).EntireColumn.Hidden = False
        End If
    Next
End Sub

Sub HideColumn()
    Sheets("Input").Range("J2:AC2").EntireColumn.Hidden = True
End Sub