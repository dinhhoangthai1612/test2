Sub CheckFilter(sheetName as String)
    Worksheets(sheetName).Select
    If Sheets(sheetName).FilterMode Then Sheets(sheetName).ShowAllData
End Sub

Sub HideRow(stRange As String, stSection As String)
    For Each cell In Range(stRange).Cells
        If cell.Value <> stSection Then
            Rows(cell.Row & ":" & cell.Row).EntireRow.Hidden = True
        End If
    Next cell
End Sub

Sub UnHideRow(stRange As String)
    For Each cell In Range(stRange).Cells
        If cell.EntireRow.Hidden = True Then
            Rows(cell.Row & ":" & cell.Row).EntireRow.Hidden = False
        End If
    Next cell
End Sub