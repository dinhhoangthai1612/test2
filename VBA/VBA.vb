Sub DisableAlert()
		Application.DisplayAlerts = False
End Sub

Sub TextID()
    Columns("A:A").NumberFormat = "@"
End Sub

Sub DeleteHidedSheet()
	Application.DisplayAlerts = False
	Dim sh As Worksheet
	For Each sh In Worksheets
	   If sh.Visible <> -1 Then
	      sh.Visible = -1
	      sh.Delete
	   End If

	   'If sh.Name = "Total" Then
            '  sh.Delete
     	   'End If
	Next

End Sub

Sub MoveHidedSheet()
    Application.DisplayAlerts = False
    Dim sh As Worksheet
    
    For Each sh In Worksheets
       If sh.Visible <> -1 Then
          sh.Move after:=Worksheets(Worksheets.Count)
       End If
    Next

End Sub

Sub ConvertText()
    Application.EnableCancelKey = xlDisabled
    Dim lastRow, i
    Dim temp As String
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 3 To lastRow
        'Col A
        temp = Left("000000", 6 - Len(Trim(Range("A" & i).Text))) & Trim(Range("A" & i).Text)
        Range("A" & i).Value = "'" & temp
        
        'Col C
        temp = Range("C" & i).Text
        Range("C" & i).Value = "'" & Replace(temp, "/", "")
        
        'Col D
        temp = Range("D" & i).Text
        Range("D" & i).Value = Left(temp, 20)
        
        'Col F
        temp = Range("F" & i).Text
        Range("F" & i).Value = "'" & Replace(temp, "/", "")
        
        'Col L
        temp = Left("0000", 4 - Len(Trim(Range("L" & i).Text))) & Trim(Range("L" & i).Text)
        Range("L" & i).Value = "'" & temp
        
        'col N
        Range("N" & i).Value = Round(Range("N" & i).Value, 2)
        
    Next i
End Sub

Sub ConvertColumnE()
    Application.EnableCancelKey = xlDisabled
    Dim lastRow, i
    Dim temp As String
    
    lastRow = Cells(Rows.Count, "E").End(xlUp).Row
    
    For i = 10 To lastRow
        'Col E
        temp = Range("E" & i).Text
        Range("E" & i).Value = "'" & temp
    Next i
End Sub

Sub FillDown()
    i = Sheets("Data").Range("B" & Rows.Count).End(xlUp).Row
    Worksheets("Data").Range("A2:A" & i).FillDown
    Worksheets("Data").Range("V2:V" & i).FillDown
    Worksheets("Data").Range("W2:W" & i).FillDown
    Worksheets("Data").Range("X2:X" & i).FillDown
End Sub

Sub FillDown1()
    i = Sheets("WINGS_QLIKVIEW").Range("A" & Rows.Count).End(xlUp).Row
    Worksheets("WINGS_QLIKVIEW").Range("F3:F" & i).FillDown
    Worksheets("WINGS_QLIKVIEW").Range("G3:G" & i).FillDown
    Worksheets("WINGS_QLIKVIEW").Range("I3:I" & i).FillDown
End Sub

Sub FillDownFA()
    i = Sheets("Data").Range("A" & Rows.Count).End(xlUp).Row
    Worksheets("Data").Range("B2:B" & i).FillDown
    Worksheets("Data").Range("C2:C" & i).FillDown
    Worksheets("Data").Range("D2:D" & i).FillDown
    Worksheets("Data").Range("E2:E" & i).FillDown
    Worksheets("Data").Range("K2:K" & i).FillDown
    Worksheets("Data").Range("L2:L" & i).FillDown
    Worksheets("Data").Range("M2:M" & i).FillDown
End Sub

Sub FillDownFA_1()
    i = Sheets("Data").Range("A" & Rows.Count).End(xlUp).Row
    Worksheets("Data").Range("B3:B" & i).FillDown
    Worksheets("Data").Range("C3:C" & i).FillDown
    Worksheets("Data").Range("D3:D" & i).FillDown
    Worksheets("Data").Range("E3:E" & i).FillDown
    Worksheets("Data").Range("K3:K" & i).FillDown
    Worksheets("Data").Range("L3:L" & i).FillDown
    Worksheets("Data").Range("M3:M" & i).FillDown
End Sub

Sub FillDownFA_2()
    i = Sheets("Data").Range("A" & Rows.Count).End(xlUp).Row
    Worksheets("Data").Range("B3:B" & i).FillDown
    Worksheets("Data").Range("C3:C" & i).FillDown
    Worksheets("Data").Range("D3:D" & i).FillDown
    Worksheets("Data").Range("E3:E" & i).FillDown
    Worksheets("Data").Range("K3:K" & i).FillDown
    Worksheets("Data").Range("L3:L" & i).FillDown
    Worksheets("Data").Range("M3:M" & i).FillDown
    Worksheets("Data").Range("N3:N" & i).FillDown
End Sub

Sub PivotField_ExpandCollapse()
	'PURPOSE: Shows how to Expand or Collapse the detail of a Pivot Field
	'SOURCE: www.TheSpreadsheetGuru.com

	Dim pf As PivotField

	Set pf = ActiveSheet.PivotTables("PivotTable1").PivotFields("Date")

	'Collapse Pivot Field
	pf.ShowDetail = False

	'Expand Pivot Field
	'pf.ShowDetail = True

End Sub

Sub Filter_PivotTable_1()
	'PURPOSE: Filter on multiple items with the Report Filter field
	'SOURCE: www.TheSpreadsheetGuru.com

	Dim pf As PivotField

	Set pf = ActiveSheet.PivotTables("PivotTable4").PivotFields("DELAY")

	'Clear Out Any Previous Filtering
	pf.ClearAllFilters

	'Enable filtering on multiple items
    pf.EnableMultiplePageItems = True
    
	'Must turn off items you do not want showing
	On Error Resume Next
	pf.PivotItems("1").Visible = False
	pf.PivotItems("2").Visible = False
	pf.PivotItems("3").Visible = False
    pf.PivotItems("4").Visible = False
	pf.PivotItems("5").Visible = False

End Sub

Sub Filter_PivotTable_2()
	'PURPOSE: Filter on multiple items with the Report Filter field
	'SOURCE: www.TheSpreadsheetGuru.com

	Dim pf As PivotField

	Set pf = ActiveSheet.PivotTables("PivotTable4").PivotFields("DELAY")

	'Clear Out Any Previous Filtering
	pf.ClearAllFilters

	'Enable filtering on multiple items
    pf.EnableMultiplePageItems = True
    
	'Must turn off items you do not want showing
    On Error Resume Next
	pf.PivotItems("2").Visible = False
	pf.PivotItems("3").Visible = False
    pf.PivotItems("4").Visible = False
	pf.PivotItems("5").Visible = False

End Sub

Sub YKKTranport_RemoveDuplicates()
    'Xác định dòng cuối trong bảng dữ liệu
    Dim DongCuoi As Long
    DongCuoi = Worksheets("Data").Cells(Rows.Count, 1).End(xlUp).Row
    'Loại bỏ giá trị trùng trong bảng, gồm cột A đến cột U
    Worksheets("Data").Range("A1:U" & DongCuoi).RemoveDuplicates Columns:=Array(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21)
End Sub

Sub CopySheetPOP()
    Dim ngay As String
    ngay = Format(DateAdd("m", -1, Now()), "yy.MM")
    If SheetNameExists(ngay) = True Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(ngay).Delete
        Sheets("Top 10 Chart_Template").Copy After:=Sheets("Monthly_Chart")
        ActiveSheet.Name = ngay
    Else
        Sheets("Top 10 Chart_Template").Copy After:=Sheets("Monthly_Chart")
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

Sub convert_unicode()
    Dim get_range, cell As Range
    Dim row_excel As Integer
    Set get_range = Worksheets(1).Range("D4:D1000")
    row_excel = 4
    
    For Each cell In get_range
        On Error Resume Next
        Cells(row_excel, 4).Value = VniToUni(cell)
        On Error GoTo 0
        row_excel = row_excel + 1
    Next cell
    
End Sub

Public Function VniToUni(str) As String
    Dim VNI, UNI, i, sUni, arrUNI() As String
    VNI = "aù,aø,aû,aõ,aï,aâ,aê,aá,aà,aå,aã,aä,aé,aè,aú,aü,aë,AÙ,AØ,AÛ,AÕ,AÏ,AÂ,AÊ,AÁ,AÀ,AÅ,AÃ,AÄ,AÉ,AÈ,AÚ,AÜ,AË,eù,eø,eû,eõ,eï,eâ,eá,eà,eå,eã,eä,EÙ,EØ,EÛ,EÕ,EÏ,EÂ,EÁ,EÀ,EÅ,EÃ,EÄ,í ,ì ,æ ,ó ,ò ,Í ,Ì ,Æ ,Ó ,Ò ,où,oø,oû,oõ,oï,oâ,ô,oá,oà,oå,oã,oä,ôù,ôø,ôû,ôõ,ôï,OÙ,OØ,OÛ,OÕ,OÏ,OÂ,Ô ,OÁ,OÀ,OÅ,OÃ,OÄ,ÔÙ,ÔØ,ÔÛ,ÔÕ,ÔÏ,uù,uø,uû,uõ,uï,ö ,öù,öø,öû,öõ,öï,UÙ,UØ,UÛ,UÕ,UÏ,Ö ,ÖÙ,ÖØ,ÖÛ,ÖÕ,ÖÏ,yù,yø,yû,yõ,î ,YÙ,YØ,YÛ,YÕ,Î ,ñ ,Ñ "
    UNI = "E1,E0,1EA3,E3,1EA1,E2,103,1EA5,1EA7,1EA9,1EAB,1EAD,1EAF,1EB1,1EB3,1EB5,1EB7,C1,C0,1EA2,C3,1EA0,C2,102,1EA4,1EA6,1EA8,1EAA,1EAC,1EAE,1EB0,1EB2,1EB4,1EB6,E9,E8,1EBB,1EBD,1EB9,EA,1EBF,1EC1,1EC3,1EC5,1EC7,C9,C8,1EBA,1EBC,1EB8,CA,1EBE,1EC0,1EC2,1EC4,1EC6,ED,EC,1EC9,129,1ECB,CD,CC,1EC8,128,1ECA,F3,F2,1ECF,F5,1ECD,F4,1A1,1ED1,1ED3,1ED5,1ED7,1ED9,1EDB,1EDD,1EDF,1EE1,1EE3,D3,D2,1ECE,D5,1ECC,D4,1A0,1ED0,1ED2,1ED4,1ED6,1ED8,1EDA,1EDC,1EDE,1EE0,1EE2,FA,F9,1EE7,169,1EE5,1B0,1EE9,1EEB,1EED,1EEF,1EF1,DA,D9,1EE6,168,1EE4,1AF,1EE8,1EEA,1EEC,1EEE,1EF0,FD,1EF3,1EF7,1EF9,1EF5,DD,1EF2,1EF6,1EF8,1EF4,111,110"
    arrUNI = Split(UNI, ",")
     For i = 1 To Len(str)
            If InStr(VNI, Mid(str, i, 2)) > 0 And Len(Mid(str, i, 2)) = 2 Then
                sUni = sUni & ChrW("&h" & arrUNI(InStr(VNI, Mid(str, i, 2)) \ 3))
                 i = i + 1
            ElseIf InStr(VNI, Mid(str, i, 1) & " ") > 0 Then
                sUni = sUni & ChrW("&h" & arrUNI(InStr(VNI, Mid(str, i, 1) & " ") \ 3))
            End If
        If InStr(VNI, Mid(str, i, 1)) = 0 Or InStr("a,A,e,E,o,O,u,U,y,Y, ", Mid(str, i, 1)) > 0 Then sUni = sUni & Mid(str, i, 1)
    Next
    VniToUni = sUni
End Function

Sub Convert_Format()
    Range("G1:G1048576").NumberFormat = "@"
    Range("I1:I1048576").NumberFormat = "@"
    Range("M1:M1048576").NumberFormat = "@"
End Sub

Sub ITEM_STRUCTURE_FORMAT()
    Columns("A:A").NumberFormat = "@"
    Columns("B:B").NumberFormat = "@"
    'Columns("H:H").NumberFormat = "@"
    Columns("I:I").NumberFormat = "@"
    
    Dim lastrow As Integer
    lastrow = Range("A" & Rows.Count).End(xlUp).Row
    For i = 3 To lastrow
        Cells(i, 1).Value = Cells(i, 1).Text
        Cells(i, 2).Value = Cells(i, 2).Text
        'Cells(i, 3).Value = Cells(i, 3).Text
        Cells(i, 9).Value = Cells(i, 9).Text
    Next i
End Sub

Sub ChangeNamePictureEAP()
    Dim shp As Shape
    Dim shpName As String
    Dim i As Integer

    i = 0

    For Each shp In ActiveSheet.Shapes
    
        If shp.Type = msoPicture Then
        
            shpName = shp.Name
            ActiveSheet.Shapes.Range(Array(shpName)).Select
            Selection.ShapeRange.Name = "Logo " & i
            i = i + 1
            
        End If
        
    Next shp
End Sub

Sub AlignPictureEAP()
    Dim shp As Shape
    Dim shpName As String
    Dim i As Integer

    i = 0

    For Each shp In ActiveSheet.Shapes
    
        If shp.Type = msoPicture Then
        
            shpName = shp.Name

            If InStr(shpName, "Picture") > 0 Then

                ActiveSheet.Shapes.Range(Array(shpName)).Select
                Selection.ShapeRange.IncrementLeft 35.52519685
                i = i + 1
            
            End If

        End If
        
    Next shp
	
	Cells.Replace What:="null", Replacement:="", LookAt:=xlWhole, _
	SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
	ReplaceFormat:=False

End Sub

Sub FilterDate(strDate As String)
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Date ")
        .PivotItems(strDate).Visible = True
    End With
End Sub

Sub InsertSignature()
    Dim DirSign, DirSignPB, MYNO, image, result, serverip As String

    Dim tblarray
	Dim arrCopy(1), arrDes(1) As String
    Dim copyFlag As Boolean
	
    Dim WS_Count As Integer
    Dim sNum As Integer
    WS_Count = ActiveWorkbook.Worksheets.Count
	Sheets("Result").Move After:=ActiveWorkbook.Sheets(WS_Count)
    For sNum = 1 To WS_Count
        If Worksheets(sNum).Name <> "Result" Then
            Worksheets(sNum).Select
			Worksheets(sNum).Range("A1").Select
            MYNO = Sheets("Result").Range("A1").Value
            DirSign = Sheets("Result").Range("A2").Value
        
            'Get Data From DATABASE
            Dim CS, RS, ConnectString, SqlString As String
        
            Set CS = CreateObject("ADODB.Connection")
            Set RS = CreateObject("ADODB.Recordset")
        
            If InStr(DirSign, "10.246.194.71") > 0 Then
                serverip = "10.246.194.1"
            Else
                serverip = "10.247.194.1"
            End If
        
            ConnectString = "Driver={ISeries Access ODBC Driver};System=" + serverip + ";Uid=vnprt;Pwd=vnprt;Library=WAVEDLIB;QueryTimeout=10"
            CS.Open (ConnectString)
        
            SqlString = "SELECT E10.MY_FORM_ID, E10.CREATE_USER,C1.USER_ID,C1.USER_NAME,E11.SIGN_IMAGE_PATH,E11.SIGN_IMAGE_FILE, PROCESS_SEQ,VARCHAR_FORMAT(E10.CREATE_DATE, 'YYYY/MM/DD HH24:mi')"
            SqlString = SqlString + " FROM AWADLIB.EAP#10 E10 "
            SqlString = SqlString + " LEFT JOIN AWADLIB.EAP#11 E11 ON E10.CREATE_USER=E11.USER_ID "
            SqlString = SqlString + " LEFT JOIN AWADLIB.CM#S01 C1 ON C1.USER_ID=E10.CREATE_USER "
            SqlString = SqlString + " WHERE MY_FORM_ID='" + MYNO + "' AND STATUS='APP' ORDER BY PROCESS_SEQ"
        
            RS.Open SqlString, CS
        
            tblarray = RS.GetRows
            Dim I As Integer
            For j = 0 To UBound(tblarray, 2)
                Set Rng = Worksheets(sNum).Range("A1:AZ500")
                If IsNull(tblarray(5, j)) And j > 0 Then
                    result = result + tblarray(3, j) + ", "
                Else
                    For Each cell In Rng
                        If cell.Value = tblarray(3, j) Then
                            I = cell.MergeArea.Cells.Count - 1
							If cell.Offset(1, 0).Value = tblarray(7, j) Then
								Select Case sNum
									Case Is = 1
										If j = 0 Then
											'MsgBox Worksheets(sNum).Cells(cell.Row - 5, cell.Column).Address
											'arrCopy(0) = "Aaa"
											arrDes(0) = Replace(Worksheets(sNum).Cells(cell.Row - 5, cell.Column).Address, "$", "")
										Else
											arrDes(1) = Replace(Worksheets(sNum).Cells(cell.Row + 1, cell.Column + I).Address, "$", "")
											copyFlag = False
										End If
									Case Is = 2
										If j = 0 Then
											arrCopy(0) = Replace(Worksheets(sNum).Cells(cell.Row - 5, cell.Column).Address, "$", "")
										Else
											arrCopy(1) = Replace(Worksheets(sNum).Cells(cell.Row + 1, cell.Column + I).Address, "$", "")
											copyFlag = True
										End If
								End Select
                            End If
							
                            If Not IsNull(tblarray(5, j)) Then
                                'Range(Replace(cell.Address, "$", "")).Select
                                'MsgBox Replace(cell.Address, "$", "")
                                If cell.Offset(1, 0).Value = tblarray(7, j) Then
                                    I = cell.MergeArea.Cells.Count - 1
                                    'MsgBox cell.Row & "____" & cell.Column & "____" & I
                                    'MsgBox Worksheets(sNum).Cells(cell.Row - 3, cell.Column).Address & "___" & Worksheets(sNum).Cells(cell.Row - 1, cell.Column + I).Address
                                    If Not (sNum = 1 And WS_Count = 3) Then
                                        Call InsertPictureInRange(DirSign + Replace(tblarray(4, j), "/", "\") + tblarray(5, j), Worksheets(sNum).Range(Worksheets(sNum).Cells(cell.Row - 3, cell.Column).Address, Worksheets(sNum).Cells(cell.Row - 1, cell.Column + I)))
                                    End If
                                End If
                            Else
                                'MsgBox cell.Value
                            End If
                        End If
                    Next cell
                End If
            Next j
        
            If InStr(result, ",") > 0 Then
                Sheets("Result").Range("A3").Value = Left(result, Len(result) - 2) + " DON'T HAVE SIGNATURE"
            End If
        
            RS.Close
            CS.Close
        End If
    Next sNum
	
	If copyFlag Then
        'MsgBox Worksheets(2).Range(arrCopy(0) & ":" & arrCopy(1)).Address
        'MsgBox Worksheets(1).Range(arrDes(0) & ":" & arrDes(1)).Address
        Worksheets(2).Range(arrCopy(0) & ":" & arrCopy(1)).Copy Destination:=Worksheets(1).Range(arrDes(0) & ":" & arrDes(1))
        Application.DisplayAlerts = False
        Worksheets(2).Delete
        Application.DisplayAlerts = True
		Worksheets(1).Select
    End If
	
End Sub

Sub InsertPictureInRange(PictureFileName As String, TargetCells As Range)
'Worksheets(sNum).Select
Dim p As Object, t As Double, l As Double, w As Double, h As Double
With TargetCells
    t = .Top
    l = .Left
    w = .Offset(0, .Columns.Count).Left - .Left
    h = .Offset(.Rows.Count, 0).Top - .Top
End With

Set p = TargetCells.Worksheet.Pictures.Insert(PictureFileName)

'MsgBox TargetCells.Worksheet.Name

With p
    .Width = w
    .Height = h
    .Top = t
    .Left = l + (w - p.Width) / 2
    .Name = PictureFileName
End With
    
Set Pic = TargetCells.Worksheet.Shapes.AddPicture(PictureFileName, _
linktofile:=msoFalse, savewithdocument:=msoCTrue, Left:=0, Top:=0, _
Width:=p.Width * 5, Height:=p.Height * 5)

With Pic
    .Height = p.Height
    .Width = p.Width
    .Top = t
    .Left = l + (w - p.Width) / 2
End With

p.Delete

End Sub

Sub CheckNumPO()

    Dim countPO As Long
    countPO = 1
    Dim StartRow, i As Long
    StartRow = 11
    
    Dim lastRow As Long
    With ActiveSheet
        lastRow = .Cells(.Rows.Count, "Z").End(xlUp).Row
    End With
    
    'Sort by Contract NO
    Range("Z10:Z" & lastRow).Select
    ActiveWorkbook.Worksheets("FORM RPA").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("FORM RPA").Sort.SortFields.Add Key:=Range("Z10"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("FORM RPA").Sort
        .SetRange Range("A11:AI" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Dim tempS As String, tempL As String
    
    Do While (StartRow <= lastRow)
        If StartRow = 11 Then
            Cells(StartRow, 1) = countPO
        Else
            tempS = Cells(StartRow, 26)
            tempL = Cells(StartRow - 1, 26)
            
            
            If tempS = tempL Then
                Cells(StartRow, 1) = countPO
            Else
                countPO = countPO + 1
                Cells(StartRow, 1) = countPO
            End If
        
        End If
        'MsgBox Cells(StartRow, 4)
        StartRow = StartRow + 1
    Loop
    
    For i = 11 To lastRow
        'col V
        temp = Range("V" & i).Text
        Range("V" & i).Value = "'" & Replace(temp, "/", "")
        'col Y
        temp = Range("Y" & i).Text
        Range("Y" & i).Value = "'" & Replace(temp, "/", "")
        'col L
        temp = Range("L" & i).Text
        Range("L" & i).Value = "'" & Replace(temp, "/", "")
        'col AA
        temp = Range("AA" & i).Text
        Range("AA" & i).Value = "'" & Replace(temp, "/", "")
    Next i
    
End Sub

Sub getDataInput()
'Dim sht As Worksheet
Dim LastRow As Long
'Set sht = ThisWorkbook.Worksheets("REMAIND PR")
LastRow = Sheets("REMAIND PR").Cells(Sheets("REMAIND PR").Rows.Count, "A").End(xlUp).Row
Dim i As Long
Dim dataInput As String, checkV As String

dataInput = "In ('"

Dim arr As Variant
ReDim arr(1 To LastRow)
arr = Sheets("REMAIND PR").Range("A1:A" & LastRow).Value

Dim temp As String
temp = "PR00000000"
'MsgBox temp
 For i = 3 To LastRow
    checkV = Replace(Trim(Worksheets("REMAIND PR").Cells(i, "A").Value), "-", "")
        If (checkV <> "" And InStr(dataInput, checkV) < 1) Then
            temp = Left(temp, Len(temp) - Len(checkV)) + checkV
            dataInput = dataInput + temp + "','"
            
        Else
        End If
 Next i
 dataInput = Left(dataInput, Len(dataInput) - 2) + ")"
 
 Sheets("REMAIND PR").Range("AZ1").Value = LastRow
 Sheets("REMAIND PR").Range("AZ2").Value = dataInput

End Sub

Sub drawBorderAndLineForAPReport(lastRow As String)
    Application.DisplayAlerts = False
    Worksheets("Report").Range("A4:L" & lastRow).Borders(xlInsideHorizontal).LineStyle = xlDash
    Worksheets("Report").Range("A4:L" & lastRow).Borders(xlInsideVertical).LineStyle = xlContinuous
    Worksheets("Report").Range("A4:L" & lastRow).BorderAround (xlContinuous)
    Worksheets("Report").Range("A:L").EntireColumn.AutoFit
End Sub

Sub drawBorderAndLineForGLAccount(lastRow As String)
    Application.DisplayAlerts = False
	Set myWorksheet = Worksheets("SCT")
	myWorksheet.Range("A8:Q" & lastRow).Font.Size = 10
    myWorksheet.Range("A8:Q" & lastRow).Borders(xlInsideHorizontal).LineStyle = xlDash
    myWorksheet.Range("A8:Q" & lastRow).Borders(xlInsideVertical).LineStyle = xlDash
    myWorksheet.Range("A8:Q" & lastRow).BorderAround (xlDouble)
    myWorksheet.Range("A:Q").EntireColumn.AutoFit
    With myWorksheet
        With .PageSetup
                .PrintArea = "A:Q"
                .PaperSize = xlPaperA4
                .Orientation = xlLandscape
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .PrintQuality = 600
                .LeftMargin = Application.InchesToPoints(0.4)
                .RightMargin = Application.InchesToPoints(0.25)
                .TopMargin = Application.InchesToPoints(0.5)
                .BottomMargin = Application.InchesToPoints(0.6)
                .HeaderMargin = Application.InchesToPoints(0.3)
                .FooterMargin = Application.InchesToPoints(0.3)
                .RightFooter = "Page &P of &N"
                .CenterHorizontally = True
                .CenterVertically = False
        End With
    End With
End Sub

Sub printYKKGlAccount()
    Application.DisplayAlerts = False
    Dim myWorksheet As Worksheet
    Set myWorksheet = Worksheets("SCT")
    myWorksheet.Range("A:Q").EntireColumn.AutoFit
    With myWorksheet
        With .PageSetup
                .PrintArea = "A:Q"
                .PaperSize = xlPaperA4
                .Orientation = xlLandscape
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .PrintQuality = 600
                '.HeaderMargin = 0.3
                '.TopMargin = 1.0
                '.RightMargin = 0.25
                '.LeftMargin = 0.45
                '.BottomMargin = 1.0
                '.FooterMargin = 0.3
                .LeftMargin = Application.InchesToPoints(0.4)
                .RightMargin = Application.InchesToPoints(0.25)
                .TopMargin = Application.InchesToPoints(0.5)
                .BottomMargin = Application.InchesToPoints(0.6)
                .HeaderMargin = Application.InchesToPoints(0.3)
                .FooterMargin = Application.InchesToPoints(0.3)
                .RightFooter = "Page &P of &N"
                .CenterHorizontally = True
                .CenterVertically = False
        End With
        .PrintOut _
            Copies:=1, _
            IgnorePrintAreas:=False
    End With
End Sub

Sub ERROR_DATA_RESCHEDULING()
    Range("A1:K1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With
    Columns("A:K").EntireColumn.AutoFit
    Columns("J:J").ColumnWidth = 80
    Columns("J:J").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Dim lastrow As Long
    lastrow = Sheets("ERROR").Range("J" & Rows.Count).End(xlUp).Row
    
    Sheets("ERROR").Activate
    
    For i = 2 To lastrow
        Range("J" & i).FormulaR1C1 = Replace(Range("J" & i).Value, "\n", Chr(10))
    Next i
	
	Sheets.Add(Before:=Sheets("ERROR")).Name = "PIVOT"
    Sheets("ERROR").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "ERROR!R1C1:R108C10", Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:="PIVOT!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion15
    Sheets("PIVOT").Select
    Cells(3, 1).Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("P.DATE")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("ORDER"), "Count of ORDER", xlCount
    ActiveSheet.PivotTables("PivotTable1").CompactLayoutRowHeader = "P.DATE"
	
End Sub

Sub ExportPictureError()
    Dim ws As Worksheet
    Dim Rng As Range
    Dim Chrt As Chart
    Dim ExportPath As String
    Dim lastRow As Long    
	
    Worksheets("PIVOT").Select
    
    lastRow = Worksheets("PIVOT").Cells(Worksheets("PIVOT").Rows.Count, "B").End(xlUp).Row
    Set Rng = Worksheets("PIVOT").Range("A3:B" & lastRow)
    
    Rng.CopyPicture
    Range("E3").PasteSpecial
    
    ActiveSheet.Shapes.AddChart(201, xlColumnClustered).Select
        
    Worksheets("PIVOT").ChartObjects(1).Width = Worksheets("PIVOT").Shapes(1).Width
    Worksheets("PIVOT").ChartObjects(1).Height = Worksheets("PIVOT").Shapes(1).Height
    
    Worksheets("PIVOT").Shapes.Range(Array(1)).Select
    Selection.Copy
    Worksheets("PIVOT").ChartObjects(1).Activate
    ActiveChart.Paste
    
    'ExportPath = ThisWorkbook.Path & "\" & Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1) & ".jpg"
    ExportPath = "H:\Robotics\Error Data Rescheduling\" & Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1) & ".jpg"
	
    Worksheets("PIVOT").ChartObjects(Worksheets("PIVOT").ChartObjects(1).Name).Activate
    ActiveChart.Export Filename:=ExportPath, Filtername:="JPG"
    
    Worksheets("PIVOT").ChartObjects(1).Delete
    Worksheets("PIVOT").Shapes(1).Delete

End Sub

Sub pasteStringtoRange(nameSheet As String, currentCell As String, index As Integer, value As String)
    Application.DisplayAlerts = False
    Sheets(nameSheet).Activate
    Dim textArray() As String, textArrayMulti() As String
    Dim temp, valueStr As String
    Dim pos, i As Integer
    i = index
    valueStr = value
    pos = InStr(valueStr, vbNewLine)
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

Function copyPrefixPicture(nameSheet As String, rangeStr1 As String, path1 As String, rangeStr2 As String, path2 As String) As Boolean
	Application.DisplayAlerts = False
	on error goto Oops
	Dim oCht As Chart
	Sheets(nameSheet).Activate
	Application.Wait (Now + TimeValue("0:00:02"))
	Range(rangeStr1).CopyPicture xlScreen, xlPicture
	set oCht =charts.add

	with oCht
		.ChartArea.Border.LineStyle = xlNone
		.PlotArea.Border.LineStyle = xlNone
		.ChartArea.Clear
		.paste
		.Export FileName:=path1, Filtername:="JPG"
	end with
	oCht.Delete
	
	Application.Wait (Now + TimeValue("0:00:02"))
	Range(rangeStr2).CopyPicture xlScreen, xlPicture
	set oCht =charts.add

	with oCht
		.ChartArea.Border.LineStyle = xlNone
		.PlotArea.Border.LineStyle = xlNone
		.ChartArea.Clear
		.paste
		.Export FileName:=path2, Filtername:="JPG"
	end with
	oCht.Delete
	copyPrefixPicture = True
	Exit Function
Oops:
	copyPrefixPicture = False
End Function

Sub SaveAsValueProdStatusReport()
    Application.DisplayAlerts = False
    Dim wsh As Worksheet
    Dim ArrayOne() As Variant
    Dim Matched As Boolean
    Dim wsName As Variant

    on error goto Oops
    ArrayOne = Array("MF","PFC","PFO","VF","ZIP TTL","CH CAP VF","Special CH","CNT CH")
    For Each wsh In ThisWorkbook.Sheets(ArrayOne)
        wsh.Cells.AutoFilter
        wsh.UsedRange.EntireColumn.AutoFit
        wsh.Cells.Copy
        wsh.Cells.PasteSpecial xlPasteValues
    Next
    Application.CutCopyMode = False

    For Each wsh In ThisWorkbook.Worksheets
        Matched = False
        For Each wsName In ArrayOne
            If wsName = wsh.Name Then
                Matched = True
                Exit For
            End If
        Next
        If Not Matched Then
            wsh.Delete
        End If
    Next wsh
Oops:
End Sub

Sub WorkbookRefreshAll()
    Application.DisplayAlerts = False
    Application.CalculateFullRebuild
    ActiveWorkbook.RefreshAll
End Sub

Sub RemoveFilterOnSheet(sheetname As String)
    Application.DisplayAlerts = False
    Sheets(sheetname).Activate
    Cells.AutoFilter
End Sub

Sub autoFitVBA(sheetname As String, rg As String)
    Application.DisplayAlerts = False
    Dim wsh As Worksheet
    Dim Matched As Boolean
    Dim wsName As String
    wsName = sheetname
    For Each wsh In ThisWorkbook.Worksheets
        Matched = False
        If wsName = wsh.Name Then
            Matched = True
        End If
        If Not Matched Then
            wsh.Delete
        End If
    Next wsh
    Sheets(sheetname).Activate
    Range(rg).EntireColumn.AutoFit
End Sub

Sub drawBorderAndLineForLiquidation(sheetname as String, rg As String, lastCol As String)
    Application.DisplayAlerts = False
    Worksheets(sheetname).Range(rg).Borders(xlInsideHorizontal).LineStyle = xlDash
    Worksheets(sheetname).Range(rg).Borders(xlInsideVertical).LineStyle = xlContinuous
    Worksheets(sheetname).Range(rg).BorderAround (xlContinuous)
    Worksheets(sheetname).Range("A:" + lastCol).EntireColumn.AutoFit
End Sub

Sub ChangeReferenceStyleA1()
    Application.ReferenceStyle = xlA1
End Sub

Sub drawBorderMonthlyMilkJob(sheetname As String, rg1 As String, rg2 As String, rg3 As String)
    Application.DisplayAlerts = False
    Worksheets(sheetname).Range(rg1).Borders(xlInsideHorizontal).LineStyle = xlDash
    Worksheets(sheetname).Range(rg1).Borders(xlInsideVertical).LineStyle = xlContinuous
    Worksheets(sheetname).Range(rg1).BorderAround (xlContinuous)
    Worksheets(sheetname).Range("A:P").EntireColumn.AutoFit
    Worksheets("Bang ki nhan").Range(rg2).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Worksheets("Bang ki nhan").Range(rg2).Borders(xlInsideVertical).LineStyle = xlContinuous
    Worksheets("Bang ki nhan").Range(rg2).BorderAround (xlContinuous)
    Worksheets("Bang ki nhan").Range("A:AI").EntireColumn.AutoFit
    Worksheets(sheetname).Range(rg3).BorderAround (xlContinuous)
End Sub

Sub SetNumberFormatRange(sheetname As String, rg As String, fm As String)
    Worksheets(sheetname).Range(rg).NumberFormat = fm
End Sub

Sub drawBorderLiquidation2(sheetname1 As String, sheetname2 As String, sheetname3 As String, rg1 As String, rg2 As String, rg3 As String)
    Application.DisplayAlerts = False
    Worksheets(sheetname1).Range(rg1).Borders(xlInsideHorizontal).LineStyle = xlDash
    Worksheets(sheetname1).Range(rg1).Borders(xlInsideVertical).LineStyle = xlContinuous
    Worksheets(sheetname1).Range(rg1).BorderAround (xlContinuous)
    Worksheets(sheetname2).Range(rg2).Borders(xlInsideHorizontal).LineStyle = xlDash
    Worksheets(sheetname2).Range(rg2).Borders(xlInsideVertical).LineStyle = xlContinuous
    Worksheets(sheetname2).Range(rg2).BorderAround (xlContinuous)
    Worksheets(sheetname3).Range(rg3).Borders(xlInsideHorizontal).LineStyle = xlDash
    Worksheets(sheetname3).Range(rg3).Borders(xlInsideVertical).LineStyle = xlContinuous
    Worksheets(sheetname3).Range(rg3).BorderAround (xlContinuous)
End Sub

Sub SetChartDataTargetBuyer(sheetName As String, rg As String)
    Application.DisplayAlerts = False
    Worksheets(sheetName).ChartObjects("Chart 7").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SetSourceData Source:=Range(rg)
End Sub

Sub SetChartSeriesTargetBuyer(sheetName As String, strCellStartDate As String, strCellEndDate As String)
    Application.DisplayAlerts = False
    Worksheets(sheetName).ChartObjects("Chart 2").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).Values = "='" & sheetName & "'!$" & strCellStartDate & "$4:$" & strCellEndDate & "$4"
    ActiveChart.SeriesCollection(1).Name = "='" & sheetName & "'!$A$4"
    ActiveChart.SeriesCollection(1).XValues = "='" & sheetName & "'!$" & strCellStartDate & "$3:$" & strCellEndDate & "$3"
    ActiveChart.SeriesCollection(2).Values = "='" & sheetName & "'!$" & strCellStartDate & "$5:$" & strCellEndDate & "$5"
    ActiveChart.SeriesCollection(2).Name = "='" & sheetName & "'!$A$5"
    ActiveChart.SeriesCollection(3).Values = "='" & sheetName & "'!$" & strCellStartDate & "$6:$" & strCellEndDate & "$6"
    ActiveChart.SeriesCollection(3).Name = "='" & sheetName & "'!$A$6"
    ActiveChart.SeriesCollection(4).Values = "='" & sheetName & "'!$" & strCellStartDate & "$7:$" & strCellEndDate & "$7"
    ActiveChart.SeriesCollection(4).Name = "='" & sheetName & "'!$A$7"
    If sheetName = "7.NIKE" Then
        ActiveChart.SeriesCollection(5).Values = "='" & sheetName & "'!$" & strCellStartDate & "$45:$" & strCellEndDate & "$45"
        ActiveChart.SeriesCollection(5).Name = "='" & sheetName & "'!$A$45"
    Else
        ActiveChart.SeriesCollection(5).Values = "='" & sheetName & "'!$" & strCellStartDate & "$39:$" & strCellEndDate & "$39"
        ActiveChart.SeriesCollection(5).Name = "='" & sheetName & "'!$A$39"
    End If
    ActiveChart.SeriesCollection(6).Values = "='" & sheetName & "'!$" & strCellStartDate & "$43:$" & strCellEndDate & "$43"
    ActiveChart.SeriesCollection(6).Name = "='" & sheetName & "'!$A$43"

    Worksheets(sheetName).ChartObjects("Chart 1").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).Values = "='" & sheetName & "'!$" & strCellStartDate & "$39:$" & strCellEndDate & "$39"
    ActiveChart.SeriesCollection(1).Name = "='" & sheetName & "'!$A$39"
    ActiveChart.SeriesCollection(1).XValues = "='" & sheetName & "'!$" & strCellStartDate & "$38:$" & strCellEndDate & "$38"
    ActiveChart.SeriesCollection(2).Values = "='" & sheetName & "'!$" & strCellStartDate & "$40:$" & strCellEndDate & "$40"
    ActiveChart.SeriesCollection(2).Name = "='" & sheetName & "'!$A$40"
    ActiveChart.SeriesCollection(3).Values = "='" & sheetName & "'!$" & strCellStartDate & "$41:$" & strCellEndDate & "$41"
    ActiveChart.SeriesCollection(3).Name = "='" & sheetName & "'!$A$41"
    ActiveChart.SeriesCollection(4).Values = "='" & sheetName & "'!$" & strCellStartDate & "$42:$" & strCellEndDate & "$42"
    ActiveChart.SeriesCollection(4).Name = "='" & sheetName & "'!$A$42"
End Sub

Sub ITEM_STRUCTRUE_RESULT()
    
    Sheets(1).Select
    Sheets(1).Copy After:=Sheets(1)
    Sheets(2).Select
    Sheets(2).Name = "RESULT"

    Sheets(2).Columns("B:I").Copy
    Sheets(2).Columns("J:J").Select
    ActiveSheet.Paste
    Range("J1:Q1").Select
    Application.CutCopyMode = False
    Selection.ClearContents

    Selection.Merge
    ActiveCell.FormulaR1C1 = "WINGS SYSTEM"
    Range("J1:Q1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
        .Font.Italic = False
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub drawBorderQuanlityInvestigate(sheetname As String, rg As String)
    Application.DisplayAlerts = False
    Worksheets(sheetname).Activate
    Range(rg).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Range(rg).Borders(xlInsideVertical).LineStyle = xlContinuous
    Range(rg).BorderAround (xlContinuous)
End Sub

Sub drawBorderAndLineForCCI(sheetname as String, rg As String)
    Application.DisplayAlerts = False
    Worksheets(sheetname).Range(rg).BorderAround Weight:=xlMedium
End Sub

Sub insertColumnDvalueReport(columnName As String)
    Application.DisplayAlerts = False
    Worksheets("Report").Activate
    Columns(columnName & ":" & columnName).Select
    Selection.Insert Shift:=xlToRight
End Sub

Sub SetChartSeriesDvalueReport(strCellEnd As String)
    Application.DisplayAlerts = False
    Dim iRow As Integer
    Dim index As Integer
    Worksheets("Report").ChartObjects("Chart 2").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).XValues = "='Report'!$D$17:$" & strCellEnd & "$17"

    index = 1
    For iRow = 18 To 33
        ActiveChart.SeriesCollection(index).Values = "='Report'!$D$" & iRow & ":$" & strCellEnd & "$" & iRow
        ActiveChart.SeriesCollection(index).Name = "='Report'!$C$" & iRow
        index = index + 1
    Next iRow
End Sub

Sub SaveChartsasImageDvalueReport(sheetName As String, chartName As String, pathImage As String)
    Application.DisplayAlerts = False
    Dim i As Integer
    Dim oCht As ChartObject

    Sheets(sheetName).Activate
    For Each ChtObj In Worksheets(sheetName).ChartObjects
        If ChtObj.Name = chartName Then
            ChtObj.Activate
            ActiveChart.Export pathImage
        End If
        If ChtObj.Name = chartName Then Exit For
    Next ChtObj
End Sub

Sub drawBorderDefectRecord(sheetname As String, rg As String)
    Application.DisplayAlerts = False
    Worksheets(sheetname).Activate
    Range(rg).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Range(rg).Borders(xlInsideVertical).LineStyle = xlContinuous
    Range(rg).BorderAround (xlContinuous)
End Sub

Sub SetChartDataDefectRecord(sheetName As String, lastRow As String)
    Application.DisplayAlerts = False
    Worksheets(sheetName).ChartObjects("Chart 1").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).Values = "='" & sheetName & "'!$V$2:$V$" & lastRow
    ActiveChart.SeriesCollection(1).Name = "='" & sheetName & "'!$V$1"
    ActiveChart.SeriesCollection(1).XValues = "='" & sheetName & "'!$U$2:$U$" & lastRow
    ActiveChart.SeriesCollection(2).Values = "='" & sheetName & "'!$W$2:$W$" & lastRow
    ActiveChart.SeriesCollection(2).Name = "='" & sheetName & "'!$W$1"
    ActiveChart.SeriesCollection(3).Values = "='" & sheetName & "'!$X$2:$X$" & lastRow
    ActiveChart.SeriesCollection(3).Name = "='" & sheetName & "'!$X$1"
    ActiveChart.SeriesCollection(4).Values = "='" & sheetName & "'!$Y$2:$Y$" & lastRow
    ActiveChart.SeriesCollection(4).Name = "='" & sheetName & "'!$Y$1"

    Worksheets(sheetName).ChartObjects("Chart 2").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).Values = "='" & sheetName & "'!$AB$2:$AB$" & lastRow
    ActiveChart.SeriesCollection(1).Name = "='" & sheetName & "'!$AB$1"
    ActiveChart.SeriesCollection(1).XValues = "='" & sheetName & "'!$AA$2:$AA$" & lastRow
    ActiveChart.SeriesCollection(2).Values = "='" & sheetName & "'!$AC$2:$AC$" & lastRow
    ActiveChart.SeriesCollection(2).Name = "='" & sheetName & "'!$AC$1"
    ActiveChart.SeriesCollection(3).Values = "='" & sheetName & "'!$AD$2:$AD$" & lastRow
    ActiveChart.SeriesCollection(3).Name = "='" & sheetName & "'!$AD$1"
    ActiveChart.SeriesCollection(4).Values = "='" & sheetName & "'!$AE$2:$AE$" & lastRow
    ActiveChart.SeriesCollection(4).Name = "='" & sheetName & "'!$AE$1"
End Sub

Sub UpdateFilterPendingAllocation()
    Application.DisplayAlerts = False
    Dim PvtTbl As PivotTable
    Set PvtTbl = Worksheets("ZIPPER").PivotTables("PivotTable3")
    PvtTbl.ClearAllFilters
    on error goto Oop1
    PvtTbl.PivotFields("CLS").PivotItems("SLD").Visible = False
    Oop1:
    on error goto Oop2
    PvtTbl.PivotFields("CLS").PivotItems("TC").Visible = False
    Oop2:
    PvtTbl.PivotFields("OR").PivotItems("(blank)").Visible = False
End Sub

Sub UpdateDailyProdRecord()
    Application.DisplayAlerts = False
    With Sheets("DATA").Activate
        Range("A5").Select
        Selection.ListObject.QueryTable.Refresh
        Application.CalculateUntilAsyncQueriesDone
    End With
    Sheets("DATA").Range("A:A,H:H,Y:Y,AB:AB,AC:AC,Q:Q,G:G,AH:AH").Copy: Sheets("DATA").Range("AK:AK").PasteSpecial Paste:=xlPasteValues
    Sheets("DATA").Range("N:N").Copy: Sheets("DATA").Range("AI:AI").PasteSpecial Paste:=xlPasteValues
    Set Rng = Range("AK2", Range("AR2").End(xlDown))
    Rng.RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7), Header:=xlYes
    Set Rng1 = Range("AI2", Range("AI2").End(xlDown))
    Rng1.RemoveDuplicates Columns:=Array(1), Header:=xlYes
End Sub

Sub RefreshOnlySheet()
    Application.DisplayAlerts = False
    With Sheets("D1_Holiday").Activate
        Range("A3").Select
        Selection.ListObject.QueryTable.Refresh
        Application.CalculateUntilAsyncQueriesDone
    End With
End Sub

Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    Dim element As Variant
    On Error GoTo IsInArrayError:
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
    Exit Function
    IsInArrayError:
    On Error GoTo 0
    IsInArray = False
End Function

Sub UpdatePivotFilterOrderArrange()
    Application.DisplayAlerts = False
    Dim pivot_item As PivotItem
    With Sheets("Pivot_Arrange").Activate
        Range("M5").Select
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("CLSCA0")
            For Each pivot_item In .PivotItems
                If IsInArray(Trim(pivot_item.Name), Array("CH", "T", "PS", "SP", "PT", "PB", "TC")) Then
                    .PivotItems(pivot_item.Name).Visible = False
                End If
            Next pivot_item
        End With
        ActiveSheet.PivotTables("PivotTable1").PivotFields("PROFLAG").ClearAllFilters
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("PROFLAG")
            on error goto Oop1
            .PivotItems("(blank)").Visible = False
            Oop1:
        End With
        Range("P5").Select
        With ActiveSheet.PivotTables("PivotTable5").PivotFields("CLSCA0")
            For Each pivot_item In .PivotItems
                If Not IsInArray(Trim(pivot_item.Name), Array("PS", "SP", "PT", "PB", "TC")) Then
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
        Range("S5").Select
        With ActiveSheet.PivotTables("PivotTable6").PivotFields("CLSCA0")
            For Each pivot_item In .PivotItems
                If Not IsInArray(Trim(pivot_item.Name), Array("CH", "T")) Then
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



Sub drawBorderSampleLT(lastRow As String)
    Application.DisplayAlerts = False
    Worksheets("Report").Range(lastRow).Borders.LineStyle = xlContinuous
End Sub

Sub insertColumnSampleLT(columnName As String)
    Application.DisplayAlerts = False
    Worksheets("Report").Activate
    Columns(columnName & ":" & columnName).Select
    Selection.Insert Shift:=xlToRight
End Sub

Sub SaveChartImageSampleLT(sheetName As String, chartName As String, pathImage As String)
    Application.DisplayAlerts = False
    Dim i As Integer
    Dim oCht As ChartObject

    Sheets(sheetName).Activate
    For Each ChtObj In Worksheets(sheetName).ChartObjects
        If ChtObj.Name = chartName Then
            ChtObj.Activate
            ActiveChart.Export pathImage
        End If
        If ChtObj.Name = chartName Then Exit For
    Next ChtObj
End Sub

Sub Replace_Special_Char()
    Sheets("ORDER").Cells.Replace What:="”", Replacement:="""", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub

Sub HideCol(TypeF As String, FromCol As String, ToCol As String)
    
    If TypeF = "HIDE" Then
        Columns(FromCol + ":" + ToCol).Hidden = True
    Else
        Columns(FromCol + ":" + ToCol).Hidden = False
    End If
End Sub


Function HideColumn_Nam(getColumn as String, sheetName as String)
    Worksheets(sheetName).Activate
    Columns(getColumn).Hidden = True
End Function




























