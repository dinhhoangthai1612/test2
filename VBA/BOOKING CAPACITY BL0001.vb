Sub pasteStringtoRange(nameSheet As String, currentCell As String, index As Integer, value As String)
    Application.DisplayAlerts = False
    Sheets(nameSheet).Activate
    Dim textArray() As String, textArrayMulti() As String
    Dim temp, valueStr As String
    Dim pos, i As Integer
    i = index
    valueStr = value
    pos = InStr(valueStr, vbNewLine)
    Worksheets(nameSheet).Select
    If pos > 0 Then
        textArrayMulti = Split(valueStr, vbNewLine)
        For Each element In textArrayMulti
            element = Trim(element)
            Do
                temp = element
                element = Replace(element, "  ", " ")
            Loop Until temp = element
            textArray = Split(element, " ")
            Set r1 = Range(currentCell & i).Resize(1, 1 + UBound(textArray))
            r1.Value = textArray
            For Each cell In r1
                cell.Formula = cell.Formula
            Next
            i = i + 1
        Next element
    Else
        valueStr = Trim(valueStr)
        Do
            temp = valueStr
            valueStr = Replace(valueStr, "  ", " ")
        Loop Until temp = valueStr
        textArray = Split(valueStr, " ")
        Set r1 = Range(currentCell & i).Resize(1, 1 + UBound(textArray))
        r1.Value = textArray
        For Each cell In r1
            cell.Formula = cell.Formula
        Next
    End If
End Sub