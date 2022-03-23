'==AWH D-VALUE (TamHB - 180121)
Sub Convert_Format()
	Range("A1:ZZ1048576").NumberFormat = "@"
End Sub


Sub MF_click_Checkbox_180121()
	ThisWorkbook.Sheets("LT").CheckBox1.Value = True
	ThisWorkbook.Sheets("LT").CheckBox2.Value = False
	ThisWorkbook.Sheets("LT").CheckBox3.Value = False
End Sub


Sub PF_click_Checkbox_180121()
	ThisWorkbook.Sheets("LT").CheckBox1.Value = False
	ThisWorkbook.Sheets("LT").CheckBox2.Value = True
	ThisWorkbook.Sheets("LT").CheckBox3.Value = False
End Sub


Sub VF_click_Checkbox_180121()
	ThisWorkbook.Sheets("LT").CheckBox1.Value = False
	ThisWorkbook.Sheets("LT").CheckBox2.Value = False
	ThisWorkbook.Sheets("LT").CheckBox3.Value = True
End Sub


Sub CommandButton1_Click()

    Range(Sheet1.Cells(5, 1), Sheet1.Cells(65536, 69)).ClearContents

    'WINGS
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim sSQL As String
    Dim Tmpstr As String
    Dim MPV As String
    
    If ThisWorkbook.Sheets("LT").CheckBox1.Value = True Then
        MPV = "'21','22'"
    End If
        
    If ThisWorkbook.Sheets("LT").CheckBox2.Value = True Then
        If MPV <> "" Then
            MPV = MPV & ","
        End If
        MPV = MPV & "'32'"
    End If
    
    If ThisWorkbook.Sheets("LT").CheckBox3.Value = True Then
        If MPV <> "" Then
            MPV = MPV & ","
        End If
        MPV = MPV & "'41','42'"
    End If

    '=== WINGS‚ÌÚ‘±Ý’è
    cn.Open "Provider=IBMDA400.DataSource.1;Data Source=" & Sheet2.Cells(2, 2) _
            & ";User ID=" & Sheet2.Cells(3, 2) _
            & ";PASSWORD=" & Sheet2.Cells(4, 2) _
            & ";Default Collection=" & Sheet2.Cells(2, 3)
    
    
    sSQL = ""
    sSQL = sSQL + "SELECT LN1C9G,LN2C9G,CN1I09,ORD1RA,PSHN9G,SMPFRA,PDSCRA,ITMC9G,IT1IA0,CLRC9G,PSHQ9G, "
    'UNIT or AM/PM
    If Sheet2.Cells(9, 2) = "1" Then
        sSQL = sSQL + "   PSDU9G,PSDT9G,WDSV02 "
    Else
        sSQL = sSQL + "   PSDU9G,PSDC9G,WDSV02 "
    End If
    sSQL = sSQL + "FROM  F9G00 LEFT JOIN C0900 ON F9G00.LN1C9G || F9G00.LN2C9G || F9G00.PDPC9G = C0900.DDTC09 AND DGRC09 = 'LN1C' "
    sSQL = sSQL + "            LEFT JOIN TRA50 ON F9G00.PSHN9G = TRA50.PSHNRA AND TRA50.EMFGRA <> '2' "
    sSQL = sSQL + "            LEFT JOIN FA000 ON F9G00.ITMC9G = FA000.ITMCA0 "
    sSQL = sSQL + "            LEFT JOIN F9H00 ON F9G00.PSHN9G = F9H00.PSHN9H AND ILRF9H = ' ' "
    sSQL = sSQL + "            LEFT JOIN C0200 ON F9G00.PSDU9G = C0200.YMDU02 "
    sSQL = sSQL + "WHERE SCTC9G IN (" & MPV & ") "
    sSQL = sSQL + "AND   PCPU9H >= '" & Sheet1.Cells(1, 4) & "' "
    sSQL = sSQL + "AND   PCPU9H <= '" & Sheet1.Cells(1, 5) & "' "
    sSQL = sSQL + "AND   DPTC02 = '01' "
    'sSQL = sSQL + "AND   PSHN9G = 'PR09233760' " ' for TEST
    sSQL = sSQL + "GROUP BY LN1C9G,LN2C9G,CN1I09,ORD1RA,PSHN9G,SMPFRA,PDSCRA,ITMC9G,IT1IA0,CLRC9G,PSHQ9G, "
    'UNIT or AM/PM
    If Sheet2.Cells(9, 2) = "1" Then
        sSQL = sSQL + "   PSDU9G,PSDT9G,WDSV02 "
    Else
        sSQL = sSQL + "   PSDU9G,PSDC9G,WDSV02 "
    End If
    sSQL = sSQL + "ORDER BY PSHN9G "
    'sSQL = sSQL + "FETCH FIRST 50 ROWS ONLY " ' for TEST
            
    rs.Open sSQL, cn
    
    m = 5
    Do Until rs.EOF
            
        For k = 0 To rs.Fields.Count - 1
        
            If k = 11 Then
                Sheet1.Cells(m, k + 1) = Mid(rs.Fields(k).Value, 1, 4) & "/" & Mid(rs.Fields(k).Value, 5, 2) & "/" & Mid(rs.Fields(k).Value, 7, 2)
            Else
                Sheet1.Cells(m, k + 1) = rs.Fields(k).Value
            End If
            
        Next
        
        
        'START
        n = 12
        If Sheet2.Cells(9, 2) = "1" Then
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
        Else
            If Sheet1.Cells(m, n + 1) = 1 Then
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & Sheet2.Cells(6, 2)
            ElseIf Sheet1.Cells(m, n + 1) = 2 Then
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & Sheet2.Cells(7, 2)
            Else
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " 0:00"
            End If
        End If
        
        'ALLOCATE DATE
        n = 16
        sSQL = ""
        sSQL = sSQL + "SELECT   RADU9G,RADT9G,WDSV02 "
        sSQL = sSQL + "FROM     F9G00 LEFT JOIN C0200 ON F9G00.RADU9G = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHN9G = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'non working day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT1
            Sheet1.Cells(m, 61) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ALLOCATE

        End If
        rs2.Close
        
        'MEASURING DATE
        n = 21
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUQ4,RADTQ4,WDSV02 "
        sSQL = sSQL + "FROM     TQ470 LEFT JOIN C0200 ON TQ470.RADUQ4 = C0200.YMDU02 "
        sSQL = sSQL + "               LEFT JOIN TRZ50 ON TQ470.CUOCQ4 = TRZ50.CUOCRZ AND TRZ50.WHDCRZ='B000' AND TRZ50.CUOCRZ <> '' " 'B000:TP/CH
        sSQL = sSQL + "WHERE    PSHNRZ = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "ORDER BY RADUQ4 DESC,RADTQ4 DESC "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'non working day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT11
            Sheet1.Cells(m, 62) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ALLOCATE

        End If
        rs2.Close
        
        'STOCK OUT DATE
        n = 26
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUV1,RADTV1,WDSV02 "
        sSQL = sSQL + "FROM     FC730V1 LEFT JOIN C0200 ON FC730V1.RADUV1 = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHNV1 = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'non working day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT12
            Sheet1.Cells(m, 63) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ALLOCATE

        End If
        rs2.Close
        
        'CHAIN DATE
        If Left(rs.Fields(0).Value, 1) = "2" Or Left(rs.Fields(0).Value, 1) = "4" Then
            'CHAIN Start
            n = 31
            If Left(rs.Fields(0).Value, 1) = "2" Then
                Tmpstr = Sheet2.Cells(12, 2)
            Else
                Tmpstr = Sheet2.Cells(13, 2)
            End If
            
            sSQL = ""
            sSQL = sSQL + "SELECT   RADUAB,RADTAB,WDSV02 "
            sSQL = sSQL + "FROM     FAB00 LEFT JOIN C0200 ON FAB00.RADUAB = C0200.YMDU02 "
            sSQL = sSQL + "WHERE    PSHNAB = '" & rs.Fields(4).Value & "' "
            sSQL = sSQL + "AND      PPTNAB IN (" & Tmpstr & ") "
            sSQL = sSQL + "AND      DPTC02 = '01' "
            sSQL = sSQL + "ORDER BY RADUAB,RADTAB "
            sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                                
            rs2.Open sSQL, cn
            If rs2.EOF = False Then
                Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
                Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
                Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
                
                'not work day
                tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
                If tmp <> 0 Then
                    Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
                Else
                    Sheet1.Cells(m, n + 4) = 0
                End If
                
                'LT15
                Sheet1.Cells(m, 64) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-CHAIN Start
            
            End If
            rs2.Close
            
            'CHAIN Finish (CP1)
            n = 36
            sSQL = ""
            sSQL = sSQL + "SELECT   CPOU9E,CPOT9E,WDSV02 "
            sSQL = sSQL + "FROM     F9E00 LEFT JOIN C0200 ON F9E00.CPOU9E = C0200.YMDU02 "
            sSQL = sSQL + "WHERE    PSHN9E = '" & rs.Fields(4).Value & "' "
            sSQL = sSQL + "AND      DPTC02 = '01' "
            sSQL = sSQL + "AND      CPSV9E = '1' "
                                
            rs2.Open sSQL, cn
            If rs2.EOF = False Then
                Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
                Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
                Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
                
                'non working day
                tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
                If tmp <> 0 Then
                    Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
                Else
                    Sheet1.Cells(m, n + 4) = 0
                End If
                
                'LT2
                Sheet1.Cells(m, 65) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-CHAIN Finish
    
            End If
            rs2.Close
        End If
        
        'ASSORT DATE
        n = 41
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUQ1,RADTQ1,WDSV02 "
        sSQL = sSQL + "FROM     TQ110 LEFT JOIN C0200 ON TQ110.RADUQ1 = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    RLTNQ1 = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "AND      BRDCQ1 = '45' "
        sSQL = sSQL + "ORDER BY RADUQ1,RADTQ1 DESC "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT3
            Sheet1.Cells(m, 66) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ASSORT
            
        End If
        rs2.Close
        
                
        'SEMI ASSEMBLE DATE
        n = 46
        If Left(rs.Fields(0).Value, 1) = "2" Then
            Tmpstr = Sheet2.Cells(16, 2)
        ElseIf Left(rs.Fields(0).Value, 1) = "3" Then
            Tmpstr = Sheet2.Cells(17, 2)
        Else
            Tmpstr = Sheet2.Cells(18, 2)
        End If
        
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUAB,RADTAB,WDSV02 "
        sSQL = sSQL + "FROM     FAB00 LEFT JOIN C0200 ON FAB00.RADUAB = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHNAB = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      PPTNAB IN (" & Tmpstr & ") "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "ORDER BY RADUAB,RADTAB "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
                                
            'LT4
            Sheet1.Cells(m, 67) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-SEMI ASSEMBLE
        
        End If
        rs2.Close
        
        
        'ASSEMBLE DATE
        n = 51
        If Left(rs.Fields(0).Value, 1) = "2" Then
            Tmpstr = Sheet2.Cells(21, 2)
        ElseIf Left(rs.Fields(0).Value, 1) = "3" Then
            Tmpstr = Sheet2.Cells(22, 2)
        Else
            Tmpstr = Sheet2.Cells(23, 2)
        End If
        
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUAB,RADTAB,WDSV02 "
        sSQL = sSQL + "FROM     FAB00 LEFT JOIN C0200 ON FAB00.RADUAB = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHNAB = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      PPTNAB IN (" & Tmpstr & ") "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "ORDER BY RADUAB,RADTAB "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
                        
            'LT5
            Sheet1.Cells(m, 68) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ASSEMBLE
        
        End If
        rs2.Close
        
        
        'COMPLETE DATE
        n = 56
        sSQL = ""
        sSQL = sSQL + "SELECT   RADU9H,RADT9H,WDSV02 "
        sSQL = sSQL + "FROM     F9H00 LEFT JOIN C0200 ON F9H00.RADU9H = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHN9H = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "AND      ILRF9H = ' ' "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
                        
            'LT6
            Sheet1.Cells(m, 69) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-COMPLETE
        
        End If
        rs2.Close
        
        
        m = m + 1
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    cn.Close

End Sub

Sub CommandMF()

    Range(Sheet1.Cells(5, 1), Sheet1.Cells(65536, 69)).ClearContents

    'WINGS
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim sSQL As String
    Dim Tmpstr As String
    Dim MPV As String
    
    
        MPV = "'21','22'"

    '=== WINGS‚ÌÚ‘±Ý’è
    cn.Open "Provider=IBMDA400.DataSource.1;Data Source=" & Sheet2.Cells(2, 2) _
            & ";User ID=" & Sheet2.Cells(3, 2) _
            & ";PASSWORD=" & Sheet2.Cells(4, 2) _
            & ";Default Collection=" & Sheet2.Cells(2, 3)
    
    
    sSQL = ""
    sSQL = sSQL + "SELECT LN1C9G,LN2C9G,CN1I09,ORD1RA,PSHN9G,SMPFRA,PDSCRA,ITMC9G,IT1IA0,CLRC9G,PSHQ9G, "
    'UNIT or AM/PM
    If Sheet2.Cells(9, 2) = "1" Then
        sSQL = sSQL + "   PSDU9G,PSDT9G,WDSV02 "
    Else
        sSQL = sSQL + "   PSDU9G,PSDC9G,WDSV02 "
    End If
    sSQL = sSQL + "FROM  F9G00 LEFT JOIN C0900 ON F9G00.LN1C9G || F9G00.LN2C9G || F9G00.PDPC9G = C0900.DDTC09 AND DGRC09 = 'LN1C' "
    sSQL = sSQL + "            LEFT JOIN TRA50 ON F9G00.PSHN9G = TRA50.PSHNRA AND TRA50.EMFGRA <> '2' "
    sSQL = sSQL + "            LEFT JOIN FA000 ON F9G00.ITMC9G = FA000.ITMCA0 "
    sSQL = sSQL + "            LEFT JOIN F9H00 ON F9G00.PSHN9G = F9H00.PSHN9H AND ILRF9H = ' ' "
    sSQL = sSQL + "            LEFT JOIN C0200 ON F9G00.PSDU9G = C0200.YMDU02 "
    sSQL = sSQL + "WHERE SCTC9G IN (" & MPV & ") "
    sSQL = sSQL + "AND   PCPU9H >= '" & Sheet1.Cells(1, 4) & "' "
    sSQL = sSQL + "AND   PCPU9H <= '" & Sheet1.Cells(1, 5) & "' "
    sSQL = sSQL + "AND   DPTC02 = '01' "
    'sSQL = sSQL + "AND   PSHN9G = 'PR09233760' " ' for TEST
    sSQL = sSQL + "GROUP BY LN1C9G,LN2C9G,CN1I09,ORD1RA,PSHN9G,SMPFRA,PDSCRA,ITMC9G,IT1IA0,CLRC9G,PSHQ9G, "
    'UNIT or AM/PM
    If Sheet2.Cells(9, 2) = "1" Then
        sSQL = sSQL + "   PSDU9G,PSDT9G,WDSV02 "
    Else
        sSQL = sSQL + "   PSDU9G,PSDC9G,WDSV02 "
    End If
    sSQL = sSQL + "ORDER BY PSHN9G "
    'sSQL = sSQL + "FETCH FIRST 50 ROWS ONLY " ' for TEST
            
    rs.Open sSQL, cn
    
    m = 5
    Do Until rs.EOF
            
        For k = 0 To rs.Fields.Count - 1
        
            If k = 11 Then
                Sheet1.Cells(m, k + 1) = Mid(rs.Fields(k).Value, 1, 4) & "/" & Mid(rs.Fields(k).Value, 5, 2) & "/" & Mid(rs.Fields(k).Value, 7, 2)
            Else
                Sheet1.Cells(m, k + 1) = rs.Fields(k).Value
            End If
            
        Next
        
        
        'START
        n = 12
        If Sheet2.Cells(9, 2) = "1" Then
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
        Else
            If Sheet1.Cells(m, n + 1) = 1 Then
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & Sheet2.Cells(6, 2)
            ElseIf Sheet1.Cells(m, n + 1) = 2 Then
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & Sheet2.Cells(7, 2)
            Else
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " 0:00"
            End If
        End If
        
        'ALLOCATE DATE
        n = 16
        sSQL = ""
        sSQL = sSQL + "SELECT   RADU9G,RADT9G,WDSV02 "
        sSQL = sSQL + "FROM     F9G00 LEFT JOIN C0200 ON F9G00.RADU9G = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHN9G = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'non working day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT1
            Sheet1.Cells(m, 61) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ALLOCATE

        End If
        rs2.Close
        
        'MEASURING DATE
        n = 21
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUQ4,RADTQ4,WDSV02 "
        sSQL = sSQL + "FROM     TQ470 LEFT JOIN C0200 ON TQ470.RADUQ4 = C0200.YMDU02 "
        sSQL = sSQL + "               LEFT JOIN TRZ50 ON TQ470.CUOCQ4 = TRZ50.CUOCRZ AND TRZ50.WHDCRZ='B000' AND TRZ50.CUOCRZ <> '' " 'B000:TP/CH
        sSQL = sSQL + "WHERE    PSHNRZ = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "ORDER BY RADUQ4 DESC,RADTQ4 DESC "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'non working day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT11
            Sheet1.Cells(m, 62) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ALLOCATE

        End If
        rs2.Close
        
        'STOCK OUT DATE
        n = 26
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUV1,RADTV1,WDSV02 "
        sSQL = sSQL + "FROM     FC730V1 LEFT JOIN C0200 ON FC730V1.RADUV1 = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHNV1 = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'non working day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT12
            Sheet1.Cells(m, 63) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ALLOCATE

        End If
        rs2.Close
        
        'CHAIN DATE
        If Left(rs.Fields(0).Value, 1) = "2" Or Left(rs.Fields(0).Value, 1) = "4" Then
            'CHAIN Start
            n = 31
            If Left(rs.Fields(0).Value, 1) = "2" Then
                Tmpstr = Sheet2.Cells(12, 2)
            Else
                Tmpstr = Sheet2.Cells(13, 2)
            End If
            
            sSQL = ""
            sSQL = sSQL + "SELECT   RADUAB,RADTAB,WDSV02 "
            sSQL = sSQL + "FROM     FAB00 LEFT JOIN C0200 ON FAB00.RADUAB = C0200.YMDU02 "
            sSQL = sSQL + "WHERE    PSHNAB = '" & rs.Fields(4).Value & "' "
            sSQL = sSQL + "AND      PPTNAB IN (" & Tmpstr & ") "
            sSQL = sSQL + "AND      DPTC02 = '01' "
            sSQL = sSQL + "ORDER BY RADUAB,RADTAB "
            sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                                
            rs2.Open sSQL, cn
            If rs2.EOF = False Then
                Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
                Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
                Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
                
                'not work day
                tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
                If tmp <> 0 Then
                    Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
                Else
                    Sheet1.Cells(m, n + 4) = 0
                End If
                
                'LT15
                Sheet1.Cells(m, 64) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-CHAIN Start
            
            End If
            rs2.Close
            
            'CHAIN Finish (CP1)
            n = 36
            sSQL = ""
            sSQL = sSQL + "SELECT   CPOU9E,CPOT9E,WDSV02 "
            sSQL = sSQL + "FROM     F9E00 LEFT JOIN C0200 ON F9E00.CPOU9E = C0200.YMDU02 "
            sSQL = sSQL + "WHERE    PSHN9E = '" & rs.Fields(4).Value & "' "
            sSQL = sSQL + "AND      DPTC02 = '01' "
            sSQL = sSQL + "AND      CPSV9E = '1' "
                                
            rs2.Open sSQL, cn
            If rs2.EOF = False Then
                Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
                Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
                Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
                
                'non working day
                tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
                If tmp <> 0 Then
                    Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
                Else
                    Sheet1.Cells(m, n + 4) = 0
                End If
                
                'LT2
                Sheet1.Cells(m, 65) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-CHAIN Finish
    
            End If
            rs2.Close
        End If
        
        'ASSORT DATE
        n = 41
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUQ1,RADTQ1,WDSV02 "
        sSQL = sSQL + "FROM     TQ110 LEFT JOIN C0200 ON TQ110.RADUQ1 = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    RLTNQ1 = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "AND      BRDCQ1 = '45' "
        sSQL = sSQL + "ORDER BY RADUQ1,RADTQ1 DESC "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT3
            Sheet1.Cells(m, 66) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ASSORT
            
        End If
        rs2.Close
        
                
        'SEMI ASSEMBLE DATE
        n = 46
        If Left(rs.Fields(0).Value, 1) = "2" Then
            Tmpstr = Sheet2.Cells(16, 2)
        ElseIf Left(rs.Fields(0).Value, 1) = "3" Then
            Tmpstr = Sheet2.Cells(17, 2)
        Else
            Tmpstr = Sheet2.Cells(18, 2)
        End If
        
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUAB,RADTAB,WDSV02 "
        sSQL = sSQL + "FROM     FAB00 LEFT JOIN C0200 ON FAB00.RADUAB = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHNAB = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      PPTNAB IN (" & Tmpstr & ") "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "ORDER BY RADUAB,RADTAB "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
                                
            'LT4
            Sheet1.Cells(m, 67) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-SEMI ASSEMBLE
        
        End If
        rs2.Close
        
        
        'ASSEMBLE DATE
        n = 51
        If Left(rs.Fields(0).Value, 1) = "2" Then
            Tmpstr = Sheet2.Cells(21, 2)
        ElseIf Left(rs.Fields(0).Value, 1) = "3" Then
            Tmpstr = Sheet2.Cells(22, 2)
        Else
            Tmpstr = Sheet2.Cells(23, 2)
        End If
        
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUAB,RADTAB,WDSV02 "
        sSQL = sSQL + "FROM     FAB00 LEFT JOIN C0200 ON FAB00.RADUAB = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHNAB = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      PPTNAB IN (" & Tmpstr & ") "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "ORDER BY RADUAB,RADTAB "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
                        
            'LT5
            Sheet1.Cells(m, 68) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ASSEMBLE
        
        End If
        rs2.Close
        
        
        'COMPLETE DATE
        n = 56
        sSQL = ""
        sSQL = sSQL + "SELECT   RADU9H,RADT9H,WDSV02 "
        sSQL = sSQL + "FROM     F9H00 LEFT JOIN C0200 ON F9H00.RADU9H = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHN9H = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "AND      ILRF9H = ' ' "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
                        
            'LT6
            Sheet1.Cells(m, 69) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-COMPLETE
        
        End If
        rs2.Close
        
        
        m = m + 1
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    cn.Close
End Sub

Sub CommandPF()


    Range(Sheet1.Cells(5, 1), Sheet1.Cells(65536, 69)).ClearContents

    'WINGS
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim sSQL As String
    Dim Tmpstr As String
    Dim MPV As String

        MPV = MPV & "'32'"
    

    '=== WINGS‚ÌÚ‘±Ý’è
    cn.Open "Provider=IBMDA400.DataSource.1;Data Source=" & Sheet2.Cells(2, 2) _
            & ";User ID=" & Sheet2.Cells(3, 2) _
            & ";PASSWORD=" & Sheet2.Cells(4, 2) _
            & ";Default Collection=" & Sheet2.Cells(2, 3)
    
    
    sSQL = ""
    sSQL = sSQL + "SELECT LN1C9G,LN2C9G,CN1I09,ORD1RA,PSHN9G,SMPFRA,PDSCRA,ITMC9G,IT1IA0,CLRC9G,PSHQ9G, "
    'UNIT or AM/PM
    If Sheet2.Cells(9, 2) = "1" Then
        sSQL = sSQL + "   PSDU9G,PSDT9G,WDSV02 "
    Else
        sSQL = sSQL + "   PSDU9G,PSDC9G,WDSV02 "
    End If
    sSQL = sSQL + "FROM  F9G00 LEFT JOIN C0900 ON F9G00.LN1C9G || F9G00.LN2C9G || F9G00.PDPC9G = C0900.DDTC09 AND DGRC09 = 'LN1C' "
    sSQL = sSQL + "            LEFT JOIN TRA50 ON F9G00.PSHN9G = TRA50.PSHNRA AND TRA50.EMFGRA <> '2' "
    sSQL = sSQL + "            LEFT JOIN FA000 ON F9G00.ITMC9G = FA000.ITMCA0 "
    sSQL = sSQL + "            LEFT JOIN F9H00 ON F9G00.PSHN9G = F9H00.PSHN9H AND ILRF9H = ' ' "
    sSQL = sSQL + "            LEFT JOIN C0200 ON F9G00.PSDU9G = C0200.YMDU02 "
    sSQL = sSQL + "WHERE SCTC9G IN (" & MPV & ") "
    sSQL = sSQL + "AND   PCPU9H >= '" & Sheet1.Cells(1, 4) & "' "
    sSQL = sSQL + "AND   PCPU9H <= '" & Sheet1.Cells(1, 5) & "' "
    sSQL = sSQL + "AND   DPTC02 = '01' "
    'sSQL = sSQL + "AND   PSHN9G = 'PR09233760' " ' for TEST
    sSQL = sSQL + "GROUP BY LN1C9G,LN2C9G,CN1I09,ORD1RA,PSHN9G,SMPFRA,PDSCRA,ITMC9G,IT1IA0,CLRC9G,PSHQ9G, "
    'UNIT or AM/PM
    If Sheet2.Cells(9, 2) = "1" Then
        sSQL = sSQL + "   PSDU9G,PSDT9G,WDSV02 "
    Else
        sSQL = sSQL + "   PSDU9G,PSDC9G,WDSV02 "
    End If
    sSQL = sSQL + "ORDER BY PSHN9G "
    'sSQL = sSQL + "FETCH FIRST 50 ROWS ONLY " ' for TEST
            
    rs.Open sSQL, cn
    
    m = 5
    Do Until rs.EOF
            
        For k = 0 To rs.Fields.Count - 1
        
            If k = 11 Then
                Sheet1.Cells(m, k + 1) = Mid(rs.Fields(k).Value, 1, 4) & "/" & Mid(rs.Fields(k).Value, 5, 2) & "/" & Mid(rs.Fields(k).Value, 7, 2)
            Else
                Sheet1.Cells(m, k + 1) = rs.Fields(k).Value
            End If
            
        Next
        
        
        'START
        n = 12
        If Sheet2.Cells(9, 2) = "1" Then
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
        Else
            If Sheet1.Cells(m, n + 1) = 1 Then
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & Sheet2.Cells(6, 2)
            ElseIf Sheet1.Cells(m, n + 1) = 2 Then
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & Sheet2.Cells(7, 2)
            Else
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " 0:00"
            End If
        End If
        
        'ALLOCATE DATE
        n = 16
        sSQL = ""
        sSQL = sSQL + "SELECT   RADU9G,RADT9G,WDSV02 "
        sSQL = sSQL + "FROM     F9G00 LEFT JOIN C0200 ON F9G00.RADU9G = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHN9G = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'non working day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT1
            Sheet1.Cells(m, 61) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ALLOCATE

        End If
        rs2.Close
        
        'MEASURING DATE
        n = 21
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUQ4,RADTQ4,WDSV02 "
        sSQL = sSQL + "FROM     TQ470 LEFT JOIN C0200 ON TQ470.RADUQ4 = C0200.YMDU02 "
        sSQL = sSQL + "               LEFT JOIN TRZ50 ON TQ470.CUOCQ4 = TRZ50.CUOCRZ AND TRZ50.WHDCRZ='B000' AND TRZ50.CUOCRZ <> '' " 'B000:TP/CH
        sSQL = sSQL + "WHERE    PSHNRZ = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "ORDER BY RADUQ4 DESC,RADTQ4 DESC "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'non working day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT11
            Sheet1.Cells(m, 62) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ALLOCATE

        End If
        rs2.Close
        
        'STOCK OUT DATE
        n = 26
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUV1,RADTV1,WDSV02 "
        sSQL = sSQL + "FROM     FC730V1 LEFT JOIN C0200 ON FC730V1.RADUV1 = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHNV1 = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'non working day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT12
            Sheet1.Cells(m, 63) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ALLOCATE

        End If
        rs2.Close
        
        'CHAIN DATE
        If Left(rs.Fields(0).Value, 1) = "2" Or Left(rs.Fields(0).Value, 1) = "4" Then
            'CHAIN Start
            n = 31
            If Left(rs.Fields(0).Value, 1) = "2" Then
                Tmpstr = Sheet2.Cells(12, 2)
            Else
                Tmpstr = Sheet2.Cells(13, 2)
            End If
            
            sSQL = ""
            sSQL = sSQL + "SELECT   RADUAB,RADTAB,WDSV02 "
            sSQL = sSQL + "FROM     FAB00 LEFT JOIN C0200 ON FAB00.RADUAB = C0200.YMDU02 "
            sSQL = sSQL + "WHERE    PSHNAB = '" & rs.Fields(4).Value & "' "
            sSQL = sSQL + "AND      PPTNAB IN (" & Tmpstr & ") "
            sSQL = sSQL + "AND      DPTC02 = '01' "
            sSQL = sSQL + "ORDER BY RADUAB,RADTAB "
            sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                                
            rs2.Open sSQL, cn
            If rs2.EOF = False Then
                Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
                Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
                Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
                
                'not work day
                tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
                If tmp <> 0 Then
                    Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
                Else
                    Sheet1.Cells(m, n + 4) = 0
                End If
                
                'LT15
                Sheet1.Cells(m, 64) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-CHAIN Start
            
            End If
            rs2.Close
            
            'CHAIN Finish (CP1)
            n = 36
            sSQL = ""
            sSQL = sSQL + "SELECT   CPOU9E,CPOT9E,WDSV02 "
            sSQL = sSQL + "FROM     F9E00 LEFT JOIN C0200 ON F9E00.CPOU9E = C0200.YMDU02 "
            sSQL = sSQL + "WHERE    PSHN9E = '" & rs.Fields(4).Value & "' "
            sSQL = sSQL + "AND      DPTC02 = '01' "
            sSQL = sSQL + "AND      CPSV9E = '1' "
                                
            rs2.Open sSQL, cn
            If rs2.EOF = False Then
                Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
                Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
                Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
                
                'non working day
                tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
                If tmp <> 0 Then
                    Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
                Else
                    Sheet1.Cells(m, n + 4) = 0
                End If
                
                'LT2
                Sheet1.Cells(m, 65) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-CHAIN Finish
    
            End If
            rs2.Close
        End If
        
        'ASSORT DATE
        n = 41
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUQ1,RADTQ1,WDSV02 "
        sSQL = sSQL + "FROM     TQ110 LEFT JOIN C0200 ON TQ110.RADUQ1 = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    RLTNQ1 = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "AND      BRDCQ1 = '45' "
        sSQL = sSQL + "ORDER BY RADUQ1,RADTQ1 DESC "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT3
            Sheet1.Cells(m, 66) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ASSORT
            
        End If
        rs2.Close
        
                
        'SEMI ASSEMBLE DATE
        n = 46
        If Left(rs.Fields(0).Value, 1) = "2" Then
            Tmpstr = Sheet2.Cells(16, 2)
        ElseIf Left(rs.Fields(0).Value, 1) = "3" Then
            Tmpstr = Sheet2.Cells(17, 2)
        Else
            Tmpstr = Sheet2.Cells(18, 2)
        End If
        
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUAB,RADTAB,WDSV02 "
        sSQL = sSQL + "FROM     FAB00 LEFT JOIN C0200 ON FAB00.RADUAB = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHNAB = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      PPTNAB IN (" & Tmpstr & ") "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "ORDER BY RADUAB,RADTAB "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
                                
            'LT4
            Sheet1.Cells(m, 67) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-SEMI ASSEMBLE
        
        End If
        rs2.Close
        
        
        'ASSEMBLE DATE
        n = 51
        If Left(rs.Fields(0).Value, 1) = "2" Then
            Tmpstr = Sheet2.Cells(21, 2)
        ElseIf Left(rs.Fields(0).Value, 1) = "3" Then
            Tmpstr = Sheet2.Cells(22, 2)
        Else
            Tmpstr = Sheet2.Cells(23, 2)
        End If
        
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUAB,RADTAB,WDSV02 "
        sSQL = sSQL + "FROM     FAB00 LEFT JOIN C0200 ON FAB00.RADUAB = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHNAB = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      PPTNAB IN (" & Tmpstr & ") "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "ORDER BY RADUAB,RADTAB "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
                        
            'LT5
            Sheet1.Cells(m, 68) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ASSEMBLE
        
        End If
        rs2.Close
        
        
        'COMPLETE DATE
        n = 56
        sSQL = ""
        sSQL = sSQL + "SELECT   RADU9H,RADT9H,WDSV02 "
        sSQL = sSQL + "FROM     F9H00 LEFT JOIN C0200 ON F9H00.RADU9H = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHN9H = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "AND      ILRF9H = ' ' "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
                        
            'LT6
            Sheet1.Cells(m, 69) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-COMPLETE
        
        End If
        rs2.Close
        
        
        m = m + 1
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    cn.Close

End Sub

Sub CommandVF()

    Range(Sheet1.Cells(5, 1), Sheet1.Cells(65536, 69)).ClearContents

    'WINGS
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim sSQL As String
    Dim Tmpstr As String
    Dim MPV As String
    

    

        MPV = MPV & "'41','42'"


    '=== WINGS‚ÌÚ‘±Ý’è
    cn.Open "Provider=IBMDA400.DataSource.1;Data Source=" & Sheet2.Cells(2, 2) _
            & ";User ID=" & Sheet2.Cells(3, 2) _
            & ";PASSWORD=" & Sheet2.Cells(4, 2) _
            & ";Default Collection=" & Sheet2.Cells(2, 3)
    
    
    sSQL = ""
    sSQL = sSQL + "SELECT LN1C9G,LN2C9G,CN1I09,ORD1RA,PSHN9G,SMPFRA,PDSCRA,ITMC9G,IT1IA0,CLRC9G,PSHQ9G, "
    'UNIT or AM/PM
    If Sheet2.Cells(9, 2) = "1" Then
        sSQL = sSQL + "   PSDU9G,PSDT9G,WDSV02 "
    Else
        sSQL = sSQL + "   PSDU9G,PSDC9G,WDSV02 "
    End If
    sSQL = sSQL + "FROM  F9G00 LEFT JOIN C0900 ON F9G00.LN1C9G || F9G00.LN2C9G || F9G00.PDPC9G = C0900.DDTC09 AND DGRC09 = 'LN1C' "
    sSQL = sSQL + "            LEFT JOIN TRA50 ON F9G00.PSHN9G = TRA50.PSHNRA AND TRA50.EMFGRA <> '2' "
    sSQL = sSQL + "            LEFT JOIN FA000 ON F9G00.ITMC9G = FA000.ITMCA0 "
    sSQL = sSQL + "            LEFT JOIN F9H00 ON F9G00.PSHN9G = F9H00.PSHN9H AND ILRF9H = ' ' "
    sSQL = sSQL + "            LEFT JOIN C0200 ON F9G00.PSDU9G = C0200.YMDU02 "
    sSQL = sSQL + "WHERE SCTC9G IN (" & MPV & ") "
    sSQL = sSQL + "AND   PCPU9H >= '" & Sheet1.Cells(1, 4) & "' "
    sSQL = sSQL + "AND   PCPU9H <= '" & Sheet1.Cells(1, 5) & "' "
    sSQL = sSQL + "AND   DPTC02 = '01' "
    'sSQL = sSQL + "AND   PSHN9G = 'PR09233760' " ' for TEST
    sSQL = sSQL + "GROUP BY LN1C9G,LN2C9G,CN1I09,ORD1RA,PSHN9G,SMPFRA,PDSCRA,ITMC9G,IT1IA0,CLRC9G,PSHQ9G, "
    'UNIT or AM/PM
    If Sheet2.Cells(9, 2) = "1" Then
        sSQL = sSQL + "   PSDU9G,PSDT9G,WDSV02 "
    Else
        sSQL = sSQL + "   PSDU9G,PSDC9G,WDSV02 "
    End If
    sSQL = sSQL + "ORDER BY PSHN9G "
    'sSQL = sSQL + "FETCH FIRST 50 ROWS ONLY " ' for TEST
            
    rs.Open sSQL, cn
    
    m = 5
    Do Until rs.EOF
            
        For k = 0 To rs.Fields.Count - 1
        
            If k = 11 Then
                Sheet1.Cells(m, k + 1) = Mid(rs.Fields(k).Value, 1, 4) & "/" & Mid(rs.Fields(k).Value, 5, 2) & "/" & Mid(rs.Fields(k).Value, 7, 2)
            Else
                Sheet1.Cells(m, k + 1) = rs.Fields(k).Value
            End If
            
        Next
        
        
        'START
        n = 12
        If Sheet2.Cells(9, 2) = "1" Then
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
        Else
            If Sheet1.Cells(m, n + 1) = 1 Then
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & Sheet2.Cells(6, 2)
            ElseIf Sheet1.Cells(m, n + 1) = 2 Then
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & Sheet2.Cells(7, 2)
            Else
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " 0:00"
            End If
        End If
        
        'ALLOCATE DATE
        n = 16
        sSQL = ""
        sSQL = sSQL + "SELECT   RADU9G,RADT9G,WDSV02 "
        sSQL = sSQL + "FROM     F9G00 LEFT JOIN C0200 ON F9G00.RADU9G = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHN9G = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'non working day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT1
            Sheet1.Cells(m, 61) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ALLOCATE

        End If
        rs2.Close
        
        'MEASURING DATE
        n = 21
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUQ4,RADTQ4,WDSV02 "
        sSQL = sSQL + "FROM     TQ470 LEFT JOIN C0200 ON TQ470.RADUQ4 = C0200.YMDU02 "
        sSQL = sSQL + "               LEFT JOIN TRZ50 ON TQ470.CUOCQ4 = TRZ50.CUOCRZ AND TRZ50.WHDCRZ='B000' AND TRZ50.CUOCRZ <> '' " 'B000:TP/CH
        sSQL = sSQL + "WHERE    PSHNRZ = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "ORDER BY RADUQ4 DESC,RADTQ4 DESC "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'non working day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT11
            Sheet1.Cells(m, 62) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ALLOCATE

        End If
        rs2.Close
        
        'STOCK OUT DATE
        n = 26
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUV1,RADTV1,WDSV02 "
        sSQL = sSQL + "FROM     FC730V1 LEFT JOIN C0200 ON FC730V1.RADUV1 = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHNV1 = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'non working day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT12
            Sheet1.Cells(m, 63) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ALLOCATE

        End If
        rs2.Close
        
        'CHAIN DATE
        If Left(rs.Fields(0).Value, 1) = "2" Or Left(rs.Fields(0).Value, 1) = "4" Then
            'CHAIN Start
            n = 31
            If Left(rs.Fields(0).Value, 1) = "2" Then
                Tmpstr = Sheet2.Cells(12, 2)
            Else
                Tmpstr = Sheet2.Cells(13, 2)
            End If
            
            sSQL = ""
            sSQL = sSQL + "SELECT   RADUAB,RADTAB,WDSV02 "
            sSQL = sSQL + "FROM     FAB00 LEFT JOIN C0200 ON FAB00.RADUAB = C0200.YMDU02 "
            sSQL = sSQL + "WHERE    PSHNAB = '" & rs.Fields(4).Value & "' "
            sSQL = sSQL + "AND      PPTNAB IN (" & Tmpstr & ") "
            sSQL = sSQL + "AND      DPTC02 = '01' "
            sSQL = sSQL + "ORDER BY RADUAB,RADTAB "
            sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                                
            rs2.Open sSQL, cn
            If rs2.EOF = False Then
                Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
                Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
                Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
                
                'not work day
                tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
                If tmp <> 0 Then
                    Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
                Else
                    Sheet1.Cells(m, n + 4) = 0
                End If
                
                'LT15
                Sheet1.Cells(m, 64) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-CHAIN Start
            
            End If
            rs2.Close
            
            'CHAIN Finish (CP1)
            n = 36
            sSQL = ""
            sSQL = sSQL + "SELECT   CPOU9E,CPOT9E,WDSV02 "
            sSQL = sSQL + "FROM     F9E00 LEFT JOIN C0200 ON F9E00.CPOU9E = C0200.YMDU02 "
            sSQL = sSQL + "WHERE    PSHN9E = '" & rs.Fields(4).Value & "' "
            sSQL = sSQL + "AND      DPTC02 = '01' "
            sSQL = sSQL + "AND      CPSV9E = '1' "
                                
            rs2.Open sSQL, cn
            If rs2.EOF = False Then
                Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
                Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
                Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
                Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
                
                'non working day
                tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
                If tmp <> 0 Then
                    Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
                Else
                    Sheet1.Cells(m, n + 4) = 0
                End If
                
                'LT2
                Sheet1.Cells(m, 65) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-CHAIN Finish
    
            End If
            rs2.Close
        End If
        
        'ASSORT DATE
        n = 41
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUQ1,RADTQ1,WDSV02 "
        sSQL = sSQL + "FROM     TQ110 LEFT JOIN C0200 ON TQ110.RADUQ1 = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    RLTNQ1 = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "AND      BRDCQ1 = '45' "
        sSQL = sSQL + "ORDER BY RADUQ1,RADTQ1 DESC "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
            
            'LT3
            Sheet1.Cells(m, 66) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ASSORT
            
        End If
        rs2.Close
        
                
        'SEMI ASSEMBLE DATE
        n = 46
        If Left(rs.Fields(0).Value, 1) = "2" Then
            Tmpstr = Sheet2.Cells(16, 2)
        ElseIf Left(rs.Fields(0).Value, 1) = "3" Then
            Tmpstr = Sheet2.Cells(17, 2)
        Else
            Tmpstr = Sheet2.Cells(18, 2)
        End If
        
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUAB,RADTAB,WDSV02 "
        sSQL = sSQL + "FROM     FAB00 LEFT JOIN C0200 ON FAB00.RADUAB = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHNAB = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      PPTNAB IN (" & Tmpstr & ") "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "ORDER BY RADUAB,RADTAB "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
                                
            'LT4
            Sheet1.Cells(m, 67) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-SEMI ASSEMBLE
        
        End If
        rs2.Close
        
        
        'ASSEMBLE DATE
        n = 51
        If Left(rs.Fields(0).Value, 1) = "2" Then
            Tmpstr = Sheet2.Cells(21, 2)
        ElseIf Left(rs.Fields(0).Value, 1) = "3" Then
            Tmpstr = Sheet2.Cells(22, 2)
        Else
            Tmpstr = Sheet2.Cells(23, 2)
        End If
        
        sSQL = ""
        sSQL = sSQL + "SELECT   RADUAB,RADTAB,WDSV02 "
        sSQL = sSQL + "FROM     FAB00 LEFT JOIN C0200 ON FAB00.RADUAB = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHNAB = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      PPTNAB IN (" & Tmpstr & ") "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "ORDER BY RADUAB,RADTAB "
        sSQL = sSQL + "FETCH FIRST 1 ROWS ONLY "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
                        
            'LT5
            Sheet1.Cells(m, 68) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-ASSEMBLE
        
        End If
        rs2.Close
        
        
        'COMPLETE DATE
        n = 56
        sSQL = ""
        sSQL = sSQL + "SELECT   RADU9H,RADT9H,WDSV02 "
        sSQL = sSQL + "FROM     F9H00 LEFT JOIN C0200 ON F9H00.RADU9H = C0200.YMDU02 "
        sSQL = sSQL + "WHERE    PSHN9H = '" & rs.Fields(4).Value & "' "
        sSQL = sSQL + "AND      DPTC02 = '01' "
        sSQL = sSQL + "AND      ILRF9H = ' ' "
                            
        rs2.Open sSQL, cn
        If rs2.EOF = False Then
            Sheet1.Cells(m, n) = Mid(rs2.Fields(0).Value, 1, 4) & "/" & Mid(rs2.Fields(0).Value, 5, 2) & "/" & Mid(rs2.Fields(0).Value, 7, 2)
            Sheet1.Cells(m, n + 1) = rs2.Fields(1).Value
            Sheet1.Cells(m, n + 2) = rs2.Fields(2).Value
            Sheet1.Cells(m, n + 3) = Sheet1.Cells(m, n) & " " & ChangeDateFormat(Sheet1.Cells(m, n + 1))
            
            'not work day
            tmp = Sheet1.Cells(m, n) - Sheet1.Cells(m, 12)
            If tmp <> 0 Then
                Sheet1.Cells(m, n + 4) = tmp - (Sheet1.Cells(m, n + 2) - Sheet1.Cells(m, 14))
            Else
                Sheet1.Cells(m, n + 4) = 0
            End If
                        
            'LT6
            Sheet1.Cells(m, 69) = ((Sheet1.Cells(m, n + 3) - Sheet1.Cells(m, 15)) * 24) - (Sheet1.Cells(m, n + 4) * 24) 'START-COMPLETE
        
        End If
        rs2.Close
        
        
        m = m + 1
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    cn.Close

End Sub

Function ChangeDateFormat(ByVal wingsdate As String) As String

    Dim newDate As String
    
    If Len(wingsdate) = 1 Then
        newDate = " 0:00:0" & Mid(wingsdate, 1, 1)
    ElseIf Len(wingsdate) = 2 Then
        newDate = " 0:00:" & Mid(wingsdate, 2, 2)
    ElseIf Len(wingsdate) = 3 Then
        newDate = " 0:0" & Mid(wingsdate, 1, 1) & ":" & Mid(wingsdate, 2, 2)
    ElseIf Len(wingsdate) = 4 Then
        newDate = " 0:" & Mid(wingsdate, 1, 2) & ":" & Mid(wingsdate, 3, 2)
    ElseIf Len(wingsdate) = 5 Then
        newDate = Mid(wingsdate, 1, 1) & ":" & Mid(wingsdate, 2, 2) & ":" & Mid(wingsdate, 4, 2)
    ElseIf Len(wingsdate) = 6 Then
        newDate = Mid(wingsdate, 1, 2) & ":" & Mid(wingsdate, 3, 2) & ":" & Mid(wingsdate, 5, 2)
    End If
    
    ChangeDateFormat = newDate

End Function