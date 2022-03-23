Sub SetRangeColorScales(index As Integer, column1 As String, column2 As String)
    Dim rg As Range
    Dim cs As ColorScale
    Dim i As Integer
    
    i = index
    Set rg = Range(column1 & "5:" & column2 & i)
    rg.FormatConditions.Delete
    'colour scale will have two colours
    Set cs = rg.FormatConditions.AddColorScale(ColorScaleType:=2)
    With cs
        'the first colour is white set at value 18
        With .ColorScaleCriteria(1)
            .FormatColor.Color = RGB(255, 255, 255)
            .Type = xlConditionValueLowestValue
        End With
        'the second colour is red
        With .ColorScaleCriteria(2)
            .FormatColor.Color = RGB(248, 105, 107)
            .Type = xlConditionValueHighestValue
        End With
    End With
End Sub