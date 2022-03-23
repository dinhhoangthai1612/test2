Sub FilterDate(strStep As String)
    On Error Resume Next

    Dim PvtItem As PivotItem
    Set Pt = Sheet1.PivotTables("PivotTable1")
    Set pf = Pt.PivotFields("Date")
    Pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
    Pt.PivotCache.Refresh
    'Select value filter Date
    For Each PvtItem In Sheet1.PivotTables("PivotTable1").PivotFields("Date").PivotItems
        If PvtItem.Visible <> True Then
            PvtItem.Visible = True
        End If    
    Next PvtItem
    'Select value filter Time
    For Each PvtItem In Sheet1.PivotTables("PivotTable1").PivotFields("Time").PivotItems
        If PvtItem.Visible <> True Then
            PvtItem.Visible = True
        End If    
    Next PvtItem

    Select Case strStep
        Case Is = "Shuttle Rack Warehouse"

            With ActiveSheet.PivotTables("PivotTable1").PivotFields("Load Size")
                .PivotItems("SR").Visible = True
            End With

        'Case Is = "Tape and Chain Warehouse"

        Case Is = "Slider Warehouse"

            With ActiveSheet.PivotTables("PivotTable1").PivotFields("Load Size")
                .PivotItems("SRS").Visible = True
            End With
    End Select
    On Error GoTo 0
End Sub