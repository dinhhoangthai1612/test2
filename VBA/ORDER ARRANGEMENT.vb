Sub UpdatePivotFilterOrderArrange()
    Application.DisplayAlerts = False
    Dim pivot_item As PivotItem
    With Sheets("Pivot_Arrange").Activate
        Range("M5").Select
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("CLSCA0")
            For Each pivot_item In .PivotItems
                If IsNumeric(Application.Match(Trim(pivot_item.Name), Array("CH", "T", "PS", "SP", "PT", "PB", "TC", "CP"), 0)) Then
                    .PivotItems(pivot_item.Name).Visible = False
                End If
            Next pivot_item
        End With
        ActiveSheet.PivotTables("PivotTable1").PivotFields("PROFLAG").ClearAllFilters
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("PROFLAG")
            On Error GoTo Oop1
            .PivotItems("(blank)").Visible = False
Oop1:
        End With
        Range("Q5").Select
        With ActiveSheet.PivotTables("PivotTable5").PivotFields("CLSCA0")
            For Each pivot_item In .PivotItems
                If Not IsNumeric(Application.Match(Trim(pivot_item.Name), Array("PS", "SP", "PT", "PB", "TC"), 0)) Then
                    .PivotItems(pivot_item.Name).Visible = False
                End If
            Next pivot_item
        End With
        ActiveSheet.PivotTables("PivotTable5").PivotFields("PROFLAG").ClearAllFilters
        With ActiveSheet.PivotTables("PivotTable5").PivotFields("PROFLAG")
            For Each pivot_item In .PivotItems
                If Trim(pivot_item.Name) <> "1" Then
                    .PivotItems(pivot_item.Name).Visible = False
                End If
            Next pivot_item
        End With
        Range("U5").Select
        With ActiveSheet.PivotTables("PivotTable6").PivotFields("CLSCA0")
            For Each pivot_item In .PivotItems
                If Not IsNumeric(Application.Match(Trim(pivot_item.Name), Array("CH", "T"), 0)) Then
                    .PivotItems(pivot_item.Name).Visible = False
                End If
            Next pivot_item
        End With
        ActiveSheet.PivotTables("PivotTable6").PivotFields("PROFLAG").ClearAllFilters
        With ActiveSheet.PivotTables("PivotTable6").PivotFields("PROFLAG")
            For Each pivot_item In .PivotItems
                If Trim(pivot_item.Name) <> "1" Then
                    .PivotItems(pivot_item.Name).Visible = False
                End If
            Next pivot_item
        End With
    End With
End Sub

Sub HideRow(stRange As String)
    For Each cell In Range(stRange).Cells
        If cell.Value < 50000 Then
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