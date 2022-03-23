Sub RemoveDups()
Sheets("AppCreateSheet").Range("A1:AG100000").RemoveDuplicates Columns:=4, Header:=xlYes
End Sub