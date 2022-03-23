Sub FilterDate(strDate As String)
    On Error Resume Next
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Date Time")
        .PivotItems(strDate).Visible = True
    End With
    On Error GoTo 0
End Sub

Sub insertDB()
    Dim rFind As Integer, strSQL As String
        Sheet1.Activate
        Sheet1.Range("A4").Select
        rFind = Rows("4:4").Find(What:="Grand Total", After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Column - 1
        
        If Not rFind = 0 Then
            'For i = 2 To rFind
                strSQL = "INSERT INTO RPALIB.RACKSTTS VALUES (TO_DATE('" & Format(Cells(4, rFind).Value, "mm/dd/yyyy hh:mm:ss") & "', 'MM/DD/YYYY HH24:MI:SS')," & _
                Cells(5, rFind).Value & "," & Cells(6, rFind).Value & "," & Cells(7, rFind).Value & "," & Cells(8, rFind).Value & ")"
                'strSQL = "INSERT INTO RPALIB.RACKSTTS VALUES (TO_DATE('" & Format(Cells(4, i).Value, "mm/dd/yyyy hh:mm:ss") & "', 'MM/DD/YYYY HH24:MI:SS')," & _
                'Cells(5, i).Value & "," & Cells(6, i).Value & "," & Cells(7, i).Value & "," & Cells(8, i).Value & ")"
                runSQL strSQL
            'Next i
        End If
     
End Sub

Sub runSQL(strSQL As String)
    Dim CS, ConnectString
    Set CS = CreateObject("ADODB.Connection")
    ConnectString = "Driver={ISeries Access ODBC Driver};System=10.246.194.1;Uid=VNPRT;Password=VNPRT;Library=WAVEDLIB;QueryTimeout=0"
    CS.Open (ConnectString)
    CS.Execute strSQL
    CS.Close
End Sub
