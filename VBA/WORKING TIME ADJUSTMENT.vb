Sub SortAscendingIDCode()
    Sheets("InOut").Select
    Range("A4:A10000").Sort Key1:=Range("A4"), Order1:=xlAscending
End Sub