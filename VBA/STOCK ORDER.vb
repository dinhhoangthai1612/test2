Sub DeleteHidedSheet()
    Application.DisplayAlerts = False
    Dim sh As Worksheet
    For Each sh In Worksheets
       If sh.Visible <> -1 Then
          sh.Visible = -1
          sh.Delete
       End If
    Next
End Sub

Sub INSERT_STOCK_ORDER()
    
    If Sheets.Count < 2 Then
        Exit Sub
    End If
    
    
    If Sheets(1).Range("L1").Value = "DONE" Then
        Exit Sub
    End If

    Dim ra As Range, strFind As String
    
    Dim requestor As String, createDate As String
    Dim salesMan As String, salesSupport As String, Cus As String, Buyer As String, Reason As String
    Dim keepCode As String, allocate As String, C_Order As String, order As String, filename As String
    
    Sheets(2).Select
    Cells(1, 1).Select
    
    '*************************Find Find Requestor*************************
    strFind = "REQUESTOR"
    requestor = ""
    createDate = ""
    
    Set ra = Sheets(2).Cells.Find(What:=strFind, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
        
    If Not ra Is Nothing Then
        For i = ra.Row + 1 To ra.Row + 10
            If Not Cells(i, ra.Column).Value = "" Then
                requestor = Cells(i, ra.Column).Value
                createDate = Cells(i + 1, ra.Column).Value
                Exit For
            End If
        Next i
        Cells(i, ra.Column).Select
    End If
    
    '*************************Find Salesman & SalesSupport*************************
    strFind = "Salesman"
    salesMan = ""
    salesSupport = ""
    
    Set ra = Sheets(2).Cells.Find(What:=strFind, After:=Range("B7"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
        
    If Not ra Is Nothing Then
        For i = ra.Column + 1 To ra.Column + 20
            If Not Cells(ra.Row, i).Value = "" And Not Cells(ra.Row, i).Value = "Sales support" Then
                If salesMan = "" Then
                    salesMan = Cells(ra.Row, i).Value
                ElseIf salesMan <> "" Then
                    salesSupport = Cells(ra.Row, i).Value
                    Exit For
                End If
            End If
        Next i
        Cells(ra.Row, i).Select
    End If
    
     '*************************Find Customer & Buyer*************************
    strFind = "Customer"
    Cus = ""
    Buyer = ""
    
    Set ra = Sheets(2).Cells.Find(What:=strFind, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
        
    If Not ra Is Nothing Then
        For i = ra.Column + 1 To ra.Column + 20
            If Not Cells(ra.Row, i).Value = "" And Not Cells(ra.Row, i).Value = "Buyer" Then
                If Cus = "" Then
                    Cus = Cells(ra.Row, i).Value
                ElseIf Cus <> "" Then
                    Buyer = Cells(ra.Row, i).Value
                    Exit For
                End If
            End If
        Next i
        Cells(ra.Row, i).Select
    End If
    
    '*************************Find Reason*************************
    strFind = "Reason need to prepare in advance:"
    Reason = ""
    
    Set ra = Sheets(2).Cells.Find(What:=strFind, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
        
    If Not ra Is Nothing Then
        For i = ra.Row + 1 To ra.Row + 10
            If Not Cells(i, ra.Column).Value = "" Then
                Reason = Cells(i, ra.Column).Value
                Exit For
            End If
        Next i
        Cells(i, ra.Column).Select
    End If
    
    '*************************Find KeepCode & Allocate*************************
    strFind = "Keep code"
    keepCode = ""
    allocate = ""
    
    Set ra = Sheets(2).Cells.Find(What:=strFind, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
        
    If Not ra Is Nothing Then
        For i = ra.Column + 1 To ra.Column + 20
            If Not Cells(ra.Row, i).Value = "" And Not Cells(ra.Row, i).Value = "Allocate keep code" Then
                If keepCode = "" Then
                    keepCode = Cells(ra.Row, i).Value
                ElseIf keepCode <> "" Then
                    allocate = Cells(ra.Row, i).Value
                    Exit For
                End If
            End If
        Next i
        Cells(ra.Row, i).Select
    End If
    
    '*************************Find C.Order No*************************
    strFind = "C.Order No"
    C_Order = ""
    
    
    Set ra = Sheets(2).Cells.Find(What:=strFind, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
        
    If Not ra Is Nothing Then
        For i = ra.Column + 1 To ra.Column + 20
            If Not Cells(ra.Row, i).Value = "" Then
                C_Order = Cells(ra.Row, i).Value
                Exit For
            End If
        Next i
        Cells(ra.Row, i).Select
    End If
        '**************************************************  INSER ORDER  **************************************************
        Sheets("Output").Select
        
        For i = 13 To 15
            If InStr(Sheets("Output").Cells(1, i).Value, "OR") = 1 Then
                INSERT_DB ThisWorkbook.Name, Sheets("Output").Cells(1, i).Text, keepCode, allocate, salesMan, salesSupport, Cus, Buyer, Reason, requestor, createDate
            End If
        Next i
        
        Sheets(1).Range("L1").Value = "DONE"
        
End Sub

Function INSERT_DB(filename As String, order As String, keep As String, allocate As String, saleMan As String, salesSup As String, Cus As String, Buyer As String, Reason As String, Req As String, req_Date As String)
    'MsgBox Application.UserName
    filename = Replace(filename, "'", "''")
    order = Replace(order, "'", "''")
    keep = Replace(keep, "'", "''")
    allocate = Replace(allocate, "'", "''")
    saleMan = Replace(saleMan, "'", "''")
    salesSup = Replace(salesSup, "'", "''")
    Cus = Replace(Cus, "'", "''")
    Buyer = Replace(Buyer, "'", "''")
    Reason = Replace(Reason, "'", "''")
    Req = Replace(Req, "'", "''")
    req_Date = Replace(req_Date, "'", "''")
    
    If Trim(req_Date) = "" Then
        req_Date = Format(Now(), "yyyy/mm/dd hh:MM")
    End If
    
    Dim strCmd, CS, ConnectString
    Set CS = CreateObject("ADODB.Connection")
    
    ConnectString = "Driver={ISeries Access ODBC Driver};System=10.246.194.1;Uid=VNPRT;Password=VNPRT;Library=WAVEDLIB;QueryTimeout=0"
            
    strCmd = "INSERT INTO RPALIB.STOCK_ORDER VALUES('" + filename + "','" + order + "','" + keep + "','" + allocate + "','" + saleMan + "'"
    strCmd = strCmd + ",'" + salesSup + "','" + Cus + "','" + Buyer + "','" + Reason + "','" + Req + "','" + req_Date + "')"
    
    Sheets(1).Range("L1").Value = strCmd
    If strCmd <> "" Then
        CS.Open (ConnectString)
        CS.Execute strCmd
        CS.Close
        
    End If
End Function


