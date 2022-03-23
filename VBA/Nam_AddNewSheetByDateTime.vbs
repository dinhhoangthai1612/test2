Sub addsheet()

Dim szToday As String

szToday = Format(Date, "yyyy-mm")

Sheets("Template").Copy After:=Sheets(ThisWorkbook.Sheets.count)

Sheets("Template (2)").name = szToday

End Sub



