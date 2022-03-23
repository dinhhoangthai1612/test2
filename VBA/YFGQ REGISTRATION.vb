Sub Unhide_ColumnsRows_On_All_Sheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
            ws.Cells.EntireColumn.Hidden = False
            ws.Cells.EntireRow.Hidden = False
    Next ws
End Sub