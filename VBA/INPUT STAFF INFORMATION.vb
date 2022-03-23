Sub Clear_Filters(sheetName As String)
    on error goto Oops
    Application.DisplayAlerts = False
    Sheets(sheetName).ShowAllData
Oops:
End Sub