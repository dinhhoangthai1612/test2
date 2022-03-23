Sub Convert_ItemCode_InQuery()
    Range("B1").NumberFormat = "@"
End Sub

Sub RemoveFilterColumns()
   ActiveSheet.ShowAllData
End Sub

Sub HideRow(index As Integer)
    Dim i As Integer
    i = index
    Sheets("Chart").Range("A" & i).EntireRow.Hidden = True
End Sub

Sub CheckFilter(sheetName as String)
    If Sheets(sheetName).FilterMode Then Sheets(sheetName).ShowAllData
End Sub