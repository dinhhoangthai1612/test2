Sub CopySheet()
    Dim ngay As String
    ngay = Format(DateAdd("y", -1, Now()), "yyyy")
    If SheetNameExists(ngay) <> True Then
        Sheets("TEMPLATE").Copy After:=Sheets("TEMPLATE")
        ActiveSheet.Name = ngay
    End If
End Sub

Function SheetNameExists(sh As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sh)
    On Error GoTo 0

    If Not ws Is Nothing Then
    SheetNameExists = True
    End If
    
End Function

Function RunCommandMF(dateFrom As String, DateTo As String)
    Range(Sheet3.Cells(5, 1), Sheet3.Cells(65536, 50)).ClearContents

    'WINGS
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim sSQL As String
    Dim m as integer
    Dim n as integer
    Dim k as integer
    Dim ret as variant
    
    '=== WINGS‚ÌÚ‘±Ý’è
    cn.Open "Provider=IBMDA400.DataSource;Data Source=" & Sheet2.Cells(2, 2) _
            & ";User ID=" & Sheet2.Cells(3, 2) _
            & ";PASSWORD=" & Sheet2.Cells(4, 2) _
            & ";Default Collection=" & Sheet2.Cells(2, 3)
    
    sSQL = ""
    sSQL = sSQL & "SELECT   LN1C9D,LN2C9D,CN1I09,PSHN9D,PSCN9D,SMPFRA,PDSCRA,ITMC9D,IT1IA0,CLRC9D,PSCQ9D,PCPQ9H,PSSU9G,PSDU9D,PSDT9D,EPFU9D,EPFT9D,"
    sSQL = sSQL & "         max(case when PPTNAB in (" & Sheet2.Cells(7, 2) & ") then RADUAB || ' ' || RADTAB end),max(case when PPTNAB in (" & Sheet2.Cells(7, 2) & ") then PPLCAB || PPMCAB end),"    'Chain
    sSQL = sSQL & "         max(case when PPTNAB in (" & Sheet2.Cells(8, 2) & ") then RADUAB || ' ' || RADTAB end),max(case when PPTNAB in (" & Sheet2.Cells(8, 2) & ") then PPLCAB || PPMCAB end),"    'Treatment
    sSQL = sSQL & "         max(case when PPTNAB in (" & Sheet2.Cells(9, 2) & ") then RADUAB || ' ' || RADTAB end),max(case when PPTNAB in (" & Sheet2.Cells(9, 2) & ") then PPLCAB || PPMCAB end),"    'Rakka
    sSQL = sSQL & "         max(case when PPTNAB in (" & Sheet2.Cells(10, 2) & ") then RADUAB || ' ' || RADTAB end),max(case when PPTNAB in (" & Sheet2.Cells(10, 2) & ") then PPLCAB || PPMCAB end),"  'Rust Preventaion
    sSQL = sSQL & "         '',"    'for Treatment summery
    sSQL = sSQL & "         max(case when PPTNAB in (" & Sheet2.Cells(11, 2) & ") then RADUAB || ' ' || RADTAB end),max(case when PPTNAB in (" & Sheet2.Cells(11, 2) & ") then PPLCAB || PPMCAB end),"  'Assembling
    sSQL = sSQL & "         PCPU9H,PCMT9H "
    sSQL = sSQL & "FROM     F9D00  LEFT JOIN C0900 ON LN1C9D || LN2C9D || PDPC9D = DDTC09 AND DGRC09 = 'LN1C' "
    sSQL = sSQL & "                LEFT JOIN F9G00 ON PSHN9D = PSHN9G "
    sSQL = sSQL & "                LEFT JOIN TRA50 ON PSHN9G = PSHNRA AND EMFGRA = '0' "
    sSQL = sSQL & "                LEFT JOIN F9H00 ON PSHN9D = PSHN9H AND ILRF9H = '' "
    sSQL = sSQL & "                LEFT JOIN FA000 ON ITMC9D = ITMCA0 "
    sSQL = sSQL & "                LEFT JOIN FAB00 ON PSHN9D = PSHNAB "
    sSQL = sSQL & "WHERE    ST1CA0 = '1' "  'ƒtƒ@ƒXƒi[
    sSQL = sSQL & "AND      ST2CA0 = '1' "  '‹à‘®
    sSQL = sSQL & "AND      ST4CA0 in ('1','2') "   '»•iorƒ`ƒF[ƒ“

    'filtering Line
    'If TextBox1.Value <> "" Then
    '    sSQL = sSQL & "AND LN1C9D = '" & TextBox1.Value & "' "
    '    If TextBox2.Value <> "" Then
    '        sSQL = sSQL & "AND LN2C9D = '" & TextBox2.Value & "' "
    '    End If
    'End If
    
    'filtering OR No.
    'If TextBox3.Value <> "" Then
    '    sSQL = sSQL & "AND PSHN9D = '" & TextBox3.Value & "' "
    'End If
    
    'EST. FINISH or ACT.FINISH
    'If OptionButton1.Value = True Then
    '    'filtering EST. FINISH
    '    If TextBox4.Value <> "" Then
    '        sSQL = sSQL & "AND EPFU9D >= '" & dateFrom & "' "
    '    End If
    '    If TextBox5.Value <> "" Then
    '        sSQL = sSQL & "AND EPFU9D <= '" & DateTo & "' "
    '    End If
    'Else
        'filtering ACT. FINISH
    '    If TextBox4.Value <> "" Then
            sSQL = sSQL & "AND PCPU9H >= '" & dateFrom & "' "
    '    End If
    '    If TextBox5.Value <> "" Then
            sSQL = sSQL & "AND PCPU9H <= '" & DateTo & "' "
    '    End If
    'End If
    
    
    sSQL = sSQL & "GROUP BY LN1C9D,LN2C9D,CN1I09,PSHN9D,PSCN9D,SMPFRA,PDSCRA,ITMC9D,IT1IA0,CLRC9D,PSCQ9D,PCPQ9H,PSSU9G,PSDU9D,PSDT9D,EPFU9D,EPFT9D,PCPU9H,PCMT9H "
            
    'EST. FINISH or ACT.FINISH
    'If OptionButton1.Value = True Then
    '    sSQL = sSQL & "ORDER BY EPFU9D,EPFT9D,PSHN9D "
    'Else
        sSQL = sSQL & "ORDER BY PCPU9H,PCMT9H,PSHN9D "
    'End If
    
    'sSQL = sSQL & "FETCH FIRST 50 ROWS ONLY "
    
    rs.Open sSQL, cn
    
    m = 5
    n = 0
    Do Until rs.EOF
        
        For k = 1 To rs.Fields.Count
                
            Sheet3.Cells(m, k) = rs.Fields(k - 1).Value
            
            If k = 18 Then
                'Chain
                ret = Application.WorksheetFunction.VLookup(Sheet3.Cells(m, 1) & Sheet3.Cells(m, 2), Sheet1.Range("Line"), 3, False)
                If IsEmpty(ret) Or Sheet3.Cells(m, 12) = 0 Then
                    Sheet3.Cells(m, k) = "-"
                End If
                    
            ElseIf k = 20 Then
                'Treatment
                ret = Application.WorksheetFunction.VLookup(Sheet3.Cells(m, 1) & Sheet3.Cells(m, 2), Sheet1.Range("Line"), 4, False)
                If IsEmpty(ret) Or Sheet3.Cells(m, 12) = 0 Then
                    Sheet3.Cells(m, k) = "-"
                End If
                
            ElseIf k = 22 Then
                'Rakka
                ret = Application.WorksheetFunction.VLookup(Sheet3.Cells(m, 1) & Sheet3.Cells(m, 2), Sheet1.Range("Line"), 5, False)
                If IsEmpty(ret) Or Sheet3.Cells(m, 12) = 0 Then
                    Sheet3.Cells(m, k) = "-"
                End If
                
            ElseIf k = 24 Then
                'Rust Preventaion
                ret = Application.WorksheetFunction.VLookup(Sheet3.Cells(m, 1) & Sheet3.Cells(m, 2), Sheet1.Range("Line"), 6, False)
                If IsEmpty(ret) Or Sheet3.Cells(m, 12) = 0 Then
                    Sheet3.Cells(m, k) = "-"
                End If
                
            ElseIf k = 26 Then
                'Treatment summery
                If Not (Sheet3.Cells(m, 21) = "" And Sheet3.Cells(m, 23) = "" And Sheet3.Cells(m, 25) = "") Then    'All nothing M/C No.
                    Sheet3.Cells(m, k) = "Z"
                End If
                
            ElseIf k = 27 Then
                'Assemble
                ret = Application.WorksheetFunction.VLookup(Sheet3.Cells(m, 1) & Sheet3.Cells(m, 2), Sheet1.Range("Line"), 7, False)
                If IsEmpty(ret) Or Sheet3.Cells(m, 12) = 0 Then
                    Sheet3.Cells(m, k) = "-"
                End If
                
            End If
            
        Next
        
        
        m = m + 1
        
        rs.MoveNext
    Loop
    
    'POP rate
    On Error Resume Next
    Sheet3.Cells(3, 18) = (Application.WorksheetFunction.CountIf(Range(Sheet3.Cells(5, 18), Sheet3.Cells(m, 18)), "<>")) / (m - 5) 'Chain
    On Error Resume Next
    Sheet3.Cells(3, 20) = (Application.WorksheetFunction.CountIf(Range(Sheet3.Cells(5, 20), Sheet3.Cells(m, 20)), "<>")) / (m - 5) 'Treatment
    On Error Resume Next
    Sheet3.Cells(3, 22) = (Application.WorksheetFunction.CountIf(Range(Sheet3.Cells(5, 22), Sheet3.Cells(m, 22)), "<>")) / (m - 5) 'Rakka
    On Error Resume Next
    Sheet3.Cells(3, 24) = (Application.WorksheetFunction.CountIf(Range(Sheet3.Cells(5, 24), Sheet3.Cells(m, 24)), "<>")) / (m - 5) 'Lust Prev.
    On Error Resume Next
    Sheet3.Cells(3, 26) = (Application.WorksheetFunction.CountIf(Range(Sheet3.Cells(5, 26), Sheet3.Cells(m, 26)), "<>")) / (m - 5) 'Treatment summery
    On Error Resume Next
    Sheet3.Cells(3, 27) = (Application.WorksheetFunction.CountIf(Range(Sheet3.Cells(5, 27), Sheet3.Cells(m, 27)), "<>")) / (m - 5) 'Assemple




    
    ' Now
    Sheet3.Cells(1, 9) = Now
    
    rs.Close
    Set rs = Nothing
    
    cn.Close
    
    'MsgBox ("UPDATE OK !!")
End Function

Function RunCommandPF(dateFrom As String, DateTo As String)

    Range(Sheet3.Cells(5, 1), Sheet3.Cells(65536, 50)).ClearContents

    'WINGS
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim sSCT As String
    Dim sSQL As String
    Dim m as integer
    Dim n as integer
    Dim k as integer
    Dim ret as variant
    
    'PF CheckBox
    'If CheckBox1.Value = True Then
        sSCT = "'32'"
    'End If
    
    'VF CheckBox
    'If CheckBox2.Value = True Then
    '    If sSCT <> "" Then
    '        sSCT = sSCT & ","
    '    End If
    '    sSCT = sSCT & "'41','42'"
    'End If

    'If sSCT = "" Then
    '    MsgBox ("Please check CheckBox.")
    'End If
    
    '=== WINGS‚ÌÚ‘±Ý’è
    cn.Open "Provider=IBMDA400.DataSource;Data Source=" & Sheet2.Cells(2, 2) _
            & ";User ID=" & Sheet2.Cells(3, 2) _
            & ";PASSWORD=" & Sheet2.Cells(4, 2) _
            & ";Default Collection=" & Sheet2.Cells(2, 3)
    
    sSQL = ""
    sSQL = sSQL & "SELECT   LN1C9D,LN2C9D,CN1I09,PSHN9D,PSCN9D,SMPFRA,PDSCRA,ITMC9D,IT1IA0,CLRC9D,PSCQ9D,PCPQ9H,PSSU9G,PSDU9D,PSDT9D,EPFU9D,EPFT9D,"
    sSQL = sSQL & "         max(case when PPTNAB in (" & Sheet2.Cells(7, 2) & ") then RADUAB || ' ' || RADTAB end),max(case when PPTNAB in (" & Sheet2.Cells(7, 2) & ") then PPLCAB || PPMCAB end),"
    sSQL = sSQL & "         max(case when PPTNAB in (" & Sheet2.Cells(8, 2) & ") then RADUAB || ' ' || RADTAB end),max(case when PPTNAB in (" & Sheet2.Cells(8, 2) & ") then PPLCAB || PPMCAB end),"
    sSQL = sSQL & "         max(case when PPTNAB in (" & Sheet2.Cells(9, 2) & ") then RADUAB || ' ' || RADTAB end),max(case when PPTNAB in (" & Sheet2.Cells(9, 2) & ") then PPLCAB || PPMCAB end),"
    sSQL = sSQL & "         PCPU9H,PCMT9H,RLTN9F "
    sSQL = sSQL & "FROM     F9D00  LEFT JOIN C0900 ON LN1C9D || LN2C9D || PDPC9D = DDTC09 AND DGRC09 = 'LN1C' "
    sSQL = sSQL & "                LEFT JOIN F9G00 ON PSHN9D = PSHN9G "
    sSQL = sSQL & "                LEFT JOIN TRA50 ON PSHN9G = PSHNRA AND EMFGRA = '0' "
    sSQL = sSQL & "                LEFT JOIN F9H00 ON PSHN9D = PSHN9H AND ILRF9H = '' "
    sSQL = sSQL & "                LEFT JOIN FA000 ON ITMC9D = ITMCA0 "
    sSQL = sSQL & "                LEFT JOIN FAB00 ON PSHN9D = PSHNAB "
    sSQL = sSQL & "                LEFT JOIN F9F00 ON PSCN9D = RLTN9F AND LCTC9F = '410' "
    sSQL = sSQL & "WHERE    SUBSTR(LN1C9D,1,2) in (" & sSCT & ") "
    
    'filtering Line
    'If TextBox1.Value <> "" Then
    '    sSQL = sSQL & "AND LN1C9D = '" & TextBox1.Value & "' "
    '    If TextBox2.Value <> "" Then
    '        sSQL = sSQL & "AND LN2C9D = '" & TextBox2.Value & "' "
    '    End If
    'End If
    
    'filtering OR No.
    'If TextBox3.Value <> "" Then
    '    sSQL = sSQL & "AND PSHN9D = '" & TextBox3.Value & "' "
    'End If
    
    'EST. FINISH or ACT.FINISH
    'If OptionButton1.Value = True Then
    '    'filtering EST. FINISH
    '    If TextBox4.Value <> "" Then
    '        sSQL = sSQL & "AND EPFU9D >= '" & TextBox4.Value & "' "
    '    End If
    '    If TextBox5.Value <> "" Then
    '        sSQL = sSQL & "AND EPFU9D <= '" & TextBox5.Value & "' "
    '    End If
    'Else
        'filtering ACT. FINISH
    '    If TextBox4.Value <> "" Then
            sSQL = sSQL & "AND PCPU9H >= '" & dateFrom & "' "
    '    End If
    '    If TextBox5.Value <> "" Then
            sSQL = sSQL & "AND PCPU9H <= '" & DateTo & "' "
    '    End If
    'End If
    
    
    sSQL = sSQL & "GROUP BY LN1C9D,LN2C9D,CN1I09,PSHN9D,PSCN9D,SMPFRA,PDSCRA,ITMC9D,IT1IA0,CLRC9D,PSCQ9D,PCPQ9H,PSSU9G,PSDU9D,PSDT9D,EPFU9D,EPFT9D,PCPU9H,PCMT9H,RLTN9F "
            
    'EST. FINISH or ACT.FINISH
    'If OptionButton1.Value = True Then
    '    sSQL = sSQL & "ORDER BY EPFU9D,EPFT9D,PSHN9D "
    'Else
        sSQL = sSQL & "ORDER BY PCPU9H,PCMT9H,PSHN9D "
    'End If
    
    'sSQL = sSQL & "FETCH FIRST 5 ROWS ONLY " 'for test
    
    rs.Open sSQL, cn
    
    m = 5
    n = 0
    Do Until rs.EOF
        
        For k = 1 To rs.Fields.Count
                
            Sheet3.Cells(m, k) = rs.Fields(k - 1).Value
            
            If k = 26 Then
                'Chain
                If Sheet3.Cells(m, k) = "" Or Sheet3.Cells(m, 12) = 0 Then
                    Sheet3.Cells(m, 18) = "-"
                End If
                    
            ElseIf k = 20 Then
                'Spacer
                If Mid(Sheet3.Cells(m, 1), 1, 2) = "41" Or Sheet3.Cells(m, 12) = 0 Then
                    Sheet3.Cells(m, k) = "-"
                End If
                
            ElseIf k = 22 Then
                'Assemble
                If Mid(Sheet3.Cells(m, 1), 1, 2) = "41" Or Sheet3.Cells(m, 12) = 0 Then
                    Sheet3.Cells(m, k) = "-"
                End If
                
            End If
            
        Next
        
        
        m = m + 1
        
        rs.MoveNext
    Loop
    
    'POP rate
    On Error Resume Next
    Sheet3.Cells(3, 18) = (Application.WorksheetFunction.CountIf(Range(Sheet3.Cells(5, 18), Sheet3.Cells(m, 18)), "<>")) / (m - 5) 'Chain
    On Error Resume Next
    Sheet3.Cells(3, 20) = (Application.WorksheetFunction.CountIf(Range(Sheet3.Cells(5, 20), Sheet3.Cells(m, 20)), "<>")) / (m - 5) 'Treatment
    On Error Resume Next
    Sheet3.Cells(3, 22) = (Application.WorksheetFunction.CountIf(Range(Sheet3.Cells(5, 22), Sheet3.Cells(m, 22)), "<>")) / (m - 5) 'Rakka
    
    ' Now
    Sheet3.Cells(1, 9) = Now
    
    rs.Close
    Set rs = Nothing
    
    cn.Close
    
    'MsgBox ("UPDATE OK !!")

End Function

Function RunCommandVF(dateFrom As String, DateTo As String)

    Range(Sheet3.Cells(5, 1), Sheet3.Cells(65536, 50)).ClearContents

    'WINGS
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim sSCT As String
    Dim sSQL As String
    Dim m as integer
    Dim n as integer
    Dim k as integer
    Dim ret as variant
    
    'PF CheckBox
    'If CheckBox1.Value = True Then
        sSCT = "'32'"
    'End If
    
    'VF CheckBox
    'If CheckBox2.Value = True Then
    '    If sSCT <> "" Then
    '        sSCT = sSCT & ","
    '    End If
        sSCT = sSCT & "'41','42'"
    'End If

    'If sSCT = "" Then
    '    MsgBox ("Please check CheckBox.")
    'End If
    
    '=== WINGS‚ÌÚ‘±Ý’è
    cn.Open "Provider=IBMDA400.DataSource;Data Source=" & Sheet2.Cells(2, 2) _
            & ";User ID=" & Sheet2.Cells(3, 2) _
            & ";PASSWORD=" & Sheet2.Cells(4, 2) _
            & ";Default Collection=" & Sheet2.Cells(2, 3)
    
    sSQL = ""
    sSQL = sSQL & "SELECT   LN1C9D,LN2C9D,CN1I09,PSHN9D,PSCN9D,SMPFRA,PDSCRA,ITMC9D,IT1IA0,CLRC9D,PSCQ9D,PCPQ9H,PSSU9G,PSDU9D,PSDT9D,EPFU9D,EPFT9D,"
    sSQL = sSQL & "         max(case when PPTNAB in (" & Sheet2.Cells(7, 2) & ") then RADUAB || ' ' || RADTAB end),max(case when PPTNAB in (" & Sheet2.Cells(7, 2) & ") then PPLCAB || PPMCAB end),"
    sSQL = sSQL & "         max(case when PPTNAB in (" & Sheet2.Cells(8, 2) & ") then RADUAB || ' ' || RADTAB end),max(case when PPTNAB in (" & Sheet2.Cells(8, 2) & ") then PPLCAB || PPMCAB end),"
    sSQL = sSQL & "         max(case when PPTNAB in (" & Sheet2.Cells(9, 2) & ") then RADUAB || ' ' || RADTAB end),max(case when PPTNAB in (" & Sheet2.Cells(9, 2) & ") then PPLCAB || PPMCAB end),"
    sSQL = sSQL & "         PCPU9H,PCMT9H,RLTN9F "
    sSQL = sSQL & "FROM     F9D00  LEFT JOIN C0900 ON LN1C9D || LN2C9D || PDPC9D = DDTC09 AND DGRC09 = 'LN1C' "
    sSQL = sSQL & "                LEFT JOIN F9G00 ON PSHN9D = PSHN9G "
    sSQL = sSQL & "                LEFT JOIN TRA50 ON PSHN9G = PSHNRA AND EMFGRA = '0' "
    sSQL = sSQL & "                LEFT JOIN F9H00 ON PSHN9D = PSHN9H AND ILRF9H = '' "
    sSQL = sSQL & "                LEFT JOIN FA000 ON ITMC9D = ITMCA0 "
    sSQL = sSQL & "                LEFT JOIN FAB00 ON PSHN9D = PSHNAB "
    sSQL = sSQL & "                LEFT JOIN F9F00 ON PSCN9D = RLTN9F AND LCTC9F = '410' "
    sSQL = sSQL & "WHERE    SUBSTR(LN1C9D,1,2) in (" & sSCT & ") "
    
    'filtering Line
    'If TextBox1.Value <> "" Then
    '    sSQL = sSQL & "AND LN1C9D = '" & TextBox1.Value & "' "
    '    If TextBox2.Value <> "" Then
    '        sSQL = sSQL & "AND LN2C9D = '" & TextBox2.Value & "' "
    '    End If
    'End If
    
    'filtering OR No.
    'If TextBox3.Value <> "" Then
    '    sSQL = sSQL & "AND PSHN9D = '" & TextBox3.Value & "' "
    'End If
    
    'EST. FINISH or ACT.FINISH
    'If OptionButton1.Value = True Then
    '    'filtering EST. FINISH
    '    If TextBox4.Value <> "" Then
    '        sSQL = sSQL & "AND EPFU9D >= '" & TextBox4.Value & "' "
    '    End If
    '    If TextBox5.Value <> "" Then
    '        sSQL = sSQL & "AND EPFU9D <= '" & TextBox5.Value & "' "
    '    End If
    'Else
        'filtering ACT. FINISH
    '    If TextBox4.Value <> "" Then
            sSQL = sSQL & "AND PCPU9H >= '" & dateFrom & "' "
    '    End If
    '    If TextBox5.Value <> "" Then
            sSQL = sSQL & "AND PCPU9H <= '" & DateTo & "' "
    '    End If
    'End If
    
    
    sSQL = sSQL & "GROUP BY LN1C9D,LN2C9D,CN1I09,PSHN9D,PSCN9D,SMPFRA,PDSCRA,ITMC9D,IT1IA0,CLRC9D,PSCQ9D,PCPQ9H,PSSU9G,PSDU9D,PSDT9D,EPFU9D,EPFT9D,PCPU9H,PCMT9H,RLTN9F "
            
    'EST. FINISH or ACT.FINISH
    'If OptionButton1.Value = True Then
    '    sSQL = sSQL & "ORDER BY EPFU9D,EPFT9D,PSHN9D "
    'Else
        sSQL = sSQL & "ORDER BY PCPU9H,PCMT9H,PSHN9D "
    'End If
    
    'sSQL = sSQL & "FETCH FIRST 5 ROWS ONLY " 'for test
    
    rs.Open sSQL, cn
    
    m = 5
    n = 0
    Do Until rs.EOF
        
        For k = 1 To rs.Fields.Count
                
            Sheet3.Cells(m, k) = rs.Fields(k - 1).Value
            
            If k = 26 Then
                'Chain
                If Sheet3.Cells(m, k) = "" Or Sheet3.Cells(m, 12) = 0 Then
                    Sheet3.Cells(m, 18) = "-"
                End If
                    
            ElseIf k = 20 Then
                'Spacer
                If Mid(Sheet3.Cells(m, 1), 1, 2) = "41" Or Sheet3.Cells(m, 12) = 0 Then
                    Sheet3.Cells(m, k) = "-"
                End If
                
            ElseIf k = 22 Then
                'Assemble
                If Mid(Sheet3.Cells(m, 1), 1, 2) = "41" Or Sheet3.Cells(m, 12) = 0 Then
                    Sheet3.Cells(m, k) = "-"
                End If
                
            End If
            
        Next
        
        
        m = m + 1
        
        rs.MoveNext
    Loop
    
    'POP rate
    On Error Resume Next
    Sheet3.Cells(3, 18) = (Application.WorksheetFunction.CountIf(Range(Sheet3.Cells(5, 18), Sheet3.Cells(m, 18)), "<>")) / (m - 5) 'Chain
    On Error Resume Next
    Sheet3.Cells(3, 20) = (Application.WorksheetFunction.CountIf(Range(Sheet3.Cells(5, 20), Sheet3.Cells(m, 20)), "<>")) / (m - 5) 'Treatment
    On Error Resume Next
    Sheet3.Cells(3, 22) = (Application.WorksheetFunction.CountIf(Range(Sheet3.Cells(5, 22), Sheet3.Cells(m, 22)), "<>")) / (m - 5) 'Rakka
    
    ' Now
    Sheet3.Cells(1, 9) = Now
    
    rs.Close
    Set rs = Nothing
    
    cn.Close
    
    'MsgBox ("UPDATE OK !!")

End Function