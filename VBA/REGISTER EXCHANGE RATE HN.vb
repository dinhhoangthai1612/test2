Sub addsheet(newSheet As String, sheetTemplate as String )

Dim szToday As String

szToday = newSheet
Sheets(sheetTemplate).Copy After:=Sheets(ThisWorkbook.Sheets.Count)

Sheets(sheetTemplate & " (2)").Name = szToday

End Sub