Sub FormulaFill()
Dim LastRowColumnA As Long
LastRowColumnA = Cells(Rows.Count, 1).End(xlUp).Row
Range("H2:H" & LastRowColumnA).Formula = "=IF(G2>0,""RSR"",""RM"")"
End Sub