Private Function FindGroupBox()
Dim i As Integer
Dim MyList As Object
Set MyList = CreateObject("System.Collections.ArrayList")
Dim strList As String
    
    For i = 1 To 20
        On Error GoTo Error
        ActiveSheet.Shapes.Range(Array("ERROR " & i)).Select
        MyList.Add "ERROR " & i
Error:
        If Err.Number = 1004 Then
          Resume Continue
        End If
        
Continue:
    Next i
    
    For i = 0 To MyList.Count - 1
        strList = strList & MyList.Item(i) & ","
    Next i
    FindGroupBox = strList
End Function


Sub SelectGroupBox(groupBox As String)
    ActiveSheet.Shapes.Range(Array(groupBox)).Select
End Sub

