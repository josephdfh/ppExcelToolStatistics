VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtLm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  On Error Resume Next
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  
  
  If sameRange(Target, shtLm.Range("clear")) Then
    Cancel = True
      
    Call clear_Lm
      
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtLm.Range("Paste_from_Pivot")) Then
    Cancel = True
      
    Call Paste_from_Pivot
      
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtLm.Range("Select_y_or_x_or_c")) Then
    Cancel = True
      
    Call toggle_select_all
      
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtLm.Range("SAS_MANOVA")) Then
    Cancel = True
    
    Call generateSAS
      
    GoTo lbl_exit
  End If
  
  'WorkingData
  If sameRange(Target, shtLm.Range("Copy_Working_Data")) Then
    Cancel = True
    
    shtLm.Range("WorkingData").Copy
    
    GoTo lbl_exit
  End If
  
lbl_exit:
  Application.EnableEvents = True
  Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  On Error GoTo lbl_exit
  Dim states(1 To 2) As Boolean
  states(1) = Application.EnableEvents
  states(2) = Application.ScreenUpdating
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  
  If Target.Row = shtLm.Range("Select_y_or_x_or_c").Row And _
     Target.Column >= shtLm.Range("DataArea_Lm[#Headers]").Cells(1, 1).Column And _
     Target.Column < shtLm.Range("DataArea_Lm[#Headers]").Cells(1, 1).Column + shtLm.Range("DataArea_Lm[#Headers]").count _
  Then
    Call order_y_c_x
    Call gen_WorkData
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtLm.Range("Ty_row")) Then
    shtLm.Range("Ty").Style = "Normal"
    shtLm.Range("Ty").Font.Color = RGB(210, 210, 210)
    With shtLm.Range("Ty").Cells(1, 1).Resize(shtLm.Range("Ty_row").Value, shtLm.Range("p").Value)
      .name = "'" & shtLm.name & "'!Ty"
      .Style = "Input"
      .Font.Bold = True
      .Font.Color = RGB(0, 0, 0)
      .HorizontalAlignment = xlCenter
    End With
    Call gen_WorkData
    
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtLm.Range("Tx_row")) Then
    shtLm.Range("Tx").Style = "Normal"
    shtLm.Range("Tx").Font.Color = RGB(210, 210, 210)
    With shtLm.Range("Tx").Cells(1, 1).Resize(shtLm.Range("Tx_row").Value, shtLm.Range("k").Value)
      .name = "'" & shtLm.name & "'!Tx"
      .Style = "Input"
      .Font.Bold = True
      .Font.Bold = True
      .Font.Color = RGB(0, 0, 0)
      .HorizontalAlignment = xlCenter
    End With
    Call gen_WorkData
    
    GoTo lbl_exit
  End If
  
  
  If Target.Row = shtLm.Range("TxData").Rows(1).Row - 2 And _
     Target.Column >= shtLm.Range("TxData").Cells(1, 1).Column And _
     Target.Column < shtLm.Range("TxData").Cells(1, 1).Column + shtLm.Range("TxData").Rows(1).Columns.count _
  Then
    Call reg
    
    GoTo lbl_exit
  End If
  
lbl_exit:
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub


Sub Paste_from_Pivot()
  Call clear_Lm
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  Dim R As Range:   Set R = shtLm.Range("Left2Data").Offset(0, 1)
  Dim lst As ListObject: Set lst = shtPivot.ListObjects("DataArea")
  
  If lst.AutoFilter.FilterMode = False Then
    lst.Range.Copy
    R.PasteSpecial Paste:=xlPasteAll
    R.ListObject.name = "DataArea_Lm"
  Else
    lst.Range.SpecialCells(xlCellTypeVisible).Copy
    R.PasteSpecial Paste:=xlPasteAll
    shtLm.ListObjects.Add(xlSrcRange, Selection, , xlYes).name = "DataArea_Lm"
  End If
  R.ListObject.DataBodyRange.Rows.Group
  With R.ListObject.HeaderRowRange.EntireRow
    .VerticalAlignment = xlTop
    .RowHeight = 30
  End With
  
  'r.Offset(0, r.ListObject.HeaderRowRange.Columns.count).Value = "Intercept"
  'shtLm.Range("DataArea_Lm[Intercept]").Value = 1
  shtLm.Range("n").Value = R.ListObject.DataBodyRange.Rows.count
  R.ListObject.HeaderRowRange.EntireColumn.HorizontalAlignment = xlCenter
  'shtLm.Range("Left2Data").Offset(0, 1).End(xlDown).Offset(2, -1).Value = "y Trans:"
  Application.EnableEvents = True
  Application.ScreenUpdating = True
End Sub


Sub clear_Lm()
  Application.EnableEvents = False
  Dim x: For Each x In shtLm.ListObjects: x.Delete: Next x:   For Each x In shtLm.ChartObjects:  x.Delete: Next x
  For Each x In Array("n", "p", "kc", "k", "Ty_row", "Tx_row", "manova_g")
    shtLm.Range(CStr(x)).ClearContents
  Next x
  
  Dim R As Range
  Set R = shtLm.Range("Left2Data").Offset(-1, 1)
  Dim last As Range: Set last = shtLm.UsedRange(shtLm.UsedRange.count)
  If last.Column >= R.Column And last.Row >= R.Row Then
    shtLm.Range(R, last).Clear
    shtLm.Range(R, last).ColumnWidth = 8.43
  Else
    R.Resize(100, 100).Clear
    R.Resize(100, 100).ColumnWidth = 8.43
  End If
  Set R = shtLm.Range("Left2Data").Offset(1000, 0).End(xlUp).Offset(1, 0)
  Set last = shtLm.UsedRange(shtLm.UsedRange.count)
  If last.Row >= R.Row Then
    shtLm.Range(R, last).EntireRow.Delete
  End If
  Set R = shtLm.Range("AAA1").End(xlToLeft).Offset(0, 5)
  Set last = shtLm.UsedRange(shtLm.UsedRange.count)
  If last.Column >= R.Column Then
    shtLm.Range(R, last).EntireColumn.Delete
  End If
  shtLm.UsedRange.ClearOutline
  Application.EnableEvents = True
End Sub

Private Sub toggle_select_all()
  If existObject("DataArea_Lm", "ListObject", shtLm) = False Then Exit Sub
  
  Dim Q(1 To 3) As Range, R As Range
  Set Q(1) = shtLm.Range("Select_y_or_x_or_c").Offset(0, 1)
  Set Q(2) = Q(1).Offset(0, shtLm.ListObjects("DataArea_Lm").HeaderRowRange.Columns.count - 1)
  For Each R In shtLm.Range(Q(1), Q(2)).Cells
    If R.Value <> "" Then
      GoTo lbl_clear
    End If
  Next R
  
'for each column, guess whether it is categorical or continuous, if continuous, rand x or y
  shtLm.Range(Q(1), Q(2)).Value = "y"
  For Each R In shtLm.Range(Q(1), Q(2))
    If TypeName(R.Offset(2, 0).Value) = "String" Then
      R.Value = "c"
      '=SUM(IF(FREQUENCY(MATCH(B7:B32,B7:B32,0),MATCH(B7:B32,B7:B32,0))>0,1))
    End If
  Next R
  GoTo lbl_exit
lbl_clear:
  shtLm.Range(Q(1), Q(2)).ClearContents
  GoTo lbl_exit
lbl_exit:
  'Call compute_Lm
End Sub

Sub generateSAS()
  On Error GoTo lbl_exit
  Dim dataname As String
  dataname = InputBox("Give datatable a name!")
  If dataname = "" Then Exit Sub
  
  Dim i As Integer
  Dim p(1 To 5) As Range, R As Range
  
  'c = categorial IV,  x = continuous IV, y = continous DV
  Dim cid(1 To 100) As Integer, ccount As Integer: ccount = 0
  Dim xid(1 To 100) As Integer, xcount As Integer: xcount = 0
  Dim yid(1 To 100) As Integer, ycount As Integer: ycount = 0
  'parse data header
  Set p(1) = shtLm.ListObjects("DataArea_Lm").HeaderRowRange
  For Each R In p(1).Cells
    If (Trim(LCase(R.Offset(-1, 0).Value)) = "c") Then
      ccount = ccount + 1
      cid(ccount) = R.Column - p(1).Cells(1, 1).Column + 1
    ElseIf (Trim(LCase(R.Offset(-1, 0).Value)) = "x") Then
      xcount = xcount + 1
      xid(xcount) = R.Column - p(1).Cells(1, 1).Column + 1
    ElseIf (Trim(LCase(R.Offset(-1, 0).Value)) = "y") Then
      ycount = ycount + 1
      yid(ycount) = R.Column - p(1).Cells(1, 1).Column + 1
    End If
  Next R
  'generate sas
  Set p(3) = shtTemp.Range("A1").Offset(0, 9999).End(xlDown).End(xlToLeft).End(xlUp).Offset(2, 0)
  Set R = p(3)
  
  R.Value = "title '" & dataname & "';"
  Set R = R.Offset(1, 0)
  R.Value = "PROC GLM DATA=" & dataname & ";"
  
'=== gen CLASS statement ====
  If ccount > 0 Then
    Set R = R.Offset(1, 0)
    R.Value = "  CLASS"
    For i = 1 To ccount
      R.Value = R.Value & " " & p(1).Cells(1, cid(i)).Value
    Next i
    R.Value = R.Value & ";"
  Else
    
  End If
  

'=== gen MODEL statement ====
  Set R = R.Offset(1, 0)
  R.Value = "  MODEL"
  For i = 1 To ycount
    R.Value = R.Value & " " & p(1).Cells(1, yid(i)).Value
  Next i
  R.Value = R.Value & " ="
  For i = 1 To ccount
    R.Value = R.Value & " " & p(1).Cells(1, cid(i)).Value
    If i < ccount Then
      R.Value = R.Value & " |"
    End If
  Next i
  If ccount > 0 And xcount > 0 Then R.Value = R.Value & " |"
  For i = 1 To xcount
    R.Value = R.Value & " " & p(1).Cells(1, xid(i)).Value
    If i < xcount Then
      R.Value = R.Value & " |"
    End If
  Next i

  R.Value = R.Value & " /ss3;   *ANOVA for each DV;"

'=== gen MANOVA statements ===
  Set R = R.Offset(1, 0)
  R.Value = "  MANOVA H ="
  For i = 1 To ccount
    R.Value = R.Value & " " & p(1).Cells(1, cid(i)).Value
    If i < ccount Then
      R.Value = R.Value & " |"
    End If
  Next i
  R.Value = R.Value & ";   *MANOVA on original DVs;"
  
  Set R = R.Offset(2, 0)
  R.Value = "  *MANOVA on differenced DVs: checking parallel profiles of every IV levels:   ;"
  Set R = R.Offset(1, 0)
  R.Value = "  MANOVA H ="
  For i = 1 To ccount
    R.Value = R.Value & " " & p(1).Cells(1, cid(i)).Value
    If i < ccount Then
      R.Value = R.Value & " |"
    End If
  Next i
  R.Value = R.Value & "   M ="
  For i = 1 To ycount - 1
    R.Value = R.Value & " " & p(1).Cells(1, yid(i)).Value & "-" & p(1).Cells(1, yid(i + 1)).Value
    If (i < ycount - 1) Then
      R.Value = R.Value & ", "
    End If
  Next i
  R.Value = R.Value & ";"
  
  Set R = R.Offset(2, 0)
  R.Value = "  *(M)ANOVA on sum of DVs: checking coincided profiles given that they are parallel:   ;"
  Set R = R.Offset(1, 0)
  R.Value = "  MANOVA H ="
  For i = 1 To ccount
    R.Value = R.Value & " " & p(1).Cells(1, cid(i)).Value
    If i < ccount Then
      R.Value = R.Value & " |"
    End If
  Next i
  R.Value = R.Value & "   M ="
  For i = 1 To ycount
    R.Value = R.Value & " " & p(1).Cells(1, yid(i)).Value
    If (i < ycount) Then
      R.Value = R.Value & " + "
    End If
  Next i
  R.Value = R.Value & ";"
  
  Set R = R.Offset(2, 0)
  R.Value = "  *Checking horizontal profiles given that ( parallel = True, coincident = False ):   ;"
  Set R = R.Offset(1, 0)
  R.Value = "  MANOVA H = INTERCEPT"
  R.Value = R.Value & "   M ="
  For i = 1 To ycount - 1
    R.Value = R.Value & " " & p(1).Cells(1, yid(i)).Value & "-" & p(1).Cells(1, yid(i + 1)).Value
    If (i < ycount - 1) Then
      R.Value = R.Value & ", "
    End If
  Next i
  R.Value = R.Value & ";"
    
  Set R = R.Offset(1, 0)
  R.Value = "QUIT;"
  
  
  
  Set R = R.Offset(2, 0)
  R.Value = "PROC GLM DATA=" & dataname & ";"
  
  Set R = R.Offset(1, 0)
  R.Value = "  CLASS"
  For i = 1 To ccount
    R.Value = R.Value & " " & p(1).Cells(1, cid(i)).Value
  Next i
  R.Value = R.Value & ";"
  
  Set R = R.Offset(1, 0)
  R.Value = "  MODEL"
  For i = 1 To ycount
    R.Value = R.Value & " " & p(1).Cells(1, yid(i)).Value
  Next i
  R.Value = R.Value & " = "
  For i = 1 To xcount
    R.Value = R.Value & " " & p(1).Cells(1, xid(i)).Value
    If i < xcount Then
      R.Value = R.Value & " |"
    End If
  Next i
  R.Value = R.Value & " /nouni;   *suppress ANOVA repitition;"
  
  
  
  Set R = R.Offset(2, 0)
  R.Value = "  *checking horizontal profiles given that ( parallel = True, coincident = True ):   ;"
  Set R = R.Offset(1, 0)
  R.Value = "  MANOVA H = INTERCEPT"
  R.Value = R.Value & "   M ="
  For i = 1 To ycount - 1
    R.Value = R.Value & " " & p(1).Cells(1, yid(i)).Value & "-" & p(1).Cells(1, yid(i + 1)).Value
    If (i < ycount - 1) Then
      R.Value = R.Value & ", "
    End If
  Next i
  R.Value = R.Value & ";"
  Set R = R.Offset(1, 0)
  R.Value = "QUIT;"
  Set R = R.Offset(1, 0)
  R.Value = "title;"
  
  SASEditor.tbox.Text = ""
  shtTemp.Range(p(3), R).Copy
  SASEditor.tbox.Paste
  Application.CutCopyMode = False
  SASEditor.Show
  
lbl_exit:
End Sub

Private Sub order_y_c_x()
  Dim states(1 To 2) As Boolean
  states(1) = Application.EnableEvents
  states(2) = Application.ScreenUpdating
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  Dim p(1 To 10) As Range, R As Range, r2 As Range, h As Range, tmp As Range
  Dim nrow As Long, ncol As Long
  
  Dim t As ListObject: Set t = shtLm.ListObjects("DataArea_Lm")
  nrow = t.DataBodyRange.Rows.count
  ncol = t.DataBodyRange.Columns.count
  Set p(1) = shtLm.Range("Left2Data").Offset(-1, 1)
  Set p(2) = shtLm.Range("Left2Data").Offset(-1, ncol)
  Set h = shtLm.Range(p(1), p(2))
  Set tmp = p(2).Offset(-1, 200).Resize(nrow + 2, 1)
  
'all blank at the end
  For Each R In h
    If LCase(Trim(R.Value)) <> "y" And LCase(Trim(R.Value)) <> "c" And LCase(Trim(R.Value)) <> "x" Then
      For Each r2 In shtLm.Range(R, p(2))
        If LCase(Trim(r2.Value)) = "y" Or LCase(Trim(r2.Value)) = "c" Or LCase(Trim(r2.Value)) = "x" Then
          'swap
          tmp.Value2 = shtLm.Range(R, R.Offset(nrow + 1, 0)).Value2
          shtLm.Range(R, R.Offset(nrow + 1, 0)).Value2 = shtLm.Range(r2, r2.Offset(nrow + 1, 0)).Value2
          shtLm.Range(r2, r2.Offset(nrow + 1, 0)).Value2 = tmp.Value2
          GoTo lbl_next_blank
        End If
      Next r2
    End If
lbl_next_blank:
  Next R
'all c after y
  For Each R In h
    If LCase(Trim(R.Value)) = "c" Then
      For Each r2 In shtLm.Range(R, p(2))
        If LCase(Trim(r2.Value)) = "y" Then
          'swap c with y
          tmp.Value2 = shtLm.Range(R, R.Offset(nrow + 1, 0)).Value2
          shtLm.Range(R, R.Offset(nrow + 1, 0)).Value2 = shtLm.Range(r2, r2.Offset(nrow + 1, 0)).Value2
          shtLm.Range(r2, r2.Offset(nrow + 1, 0)).Value2 = tmp.Value2
          GoTo lbl_next_c
        End If
      Next r2
    End If
lbl_next_c:
  Next R
  
'all x after y
  For Each R In h
    If LCase(Trim(R.Value)) = "x" Then
      For Each r2 In shtLm.Range(R, p(2))
        If LCase(Trim(r2.Value)) = "y" Then
          'swap x with y
          tmp.Value2 = shtLm.Range(R, R.Offset(nrow + 1, 0)).Value2
          shtLm.Range(R, R.Offset(nrow + 1, 0)).Value2 = shtLm.Range(r2, r2.Offset(nrow + 1, 0)).Value2
          shtLm.Range(r2, r2.Offset(nrow + 1, 0)).Value2 = tmp.Value2
          GoTo lbl_next_x
        End If
      Next r2
    End If
lbl_next_x:
  Next R

'all x after c
  For Each R In h
    If LCase(Trim(R.Value)) = "x" Then
      For Each r2 In shtLm.Range(R, p(2))
        If LCase(Trim(r2.Value)) = "c" Then
          'swap c with x
          tmp.Value2 = shtLm.Range(R, R.Offset(nrow + 1, 0)).Value2
          shtLm.Range(R, R.Offset(nrow + 1, 0)).Value2 = shtLm.Range(r2, r2.Offset(nrow + 1, 0)).Value2
          shtLm.Range(r2, r2.Offset(nrow + 1, 0)).Value2 = tmp.Value2
          GoTo lbl_next_xx
        End If
      Next r2
    End If
lbl_next_xx:
  Next R
  
  Dim ycount As Integer, ccount As Integer, xcount As Integer, i As Integer
  ycount = 0: ccount = 0: xcount = 0
  For Each R In shtLm.Range(p(1), p(2))
    Select Case LCase(Trim(R.Value))
      Case "y": ycount = ycount + 1
      Case "c": ccount = ccount + 1
      Case "x": xcount = xcount + 1
    End Select
  Next R
  shtLm.Range("p").Value = ycount
  shtLm.Range("kc").Value = ccount
  shtLm.Range("k").Value = xcount
    
  shtLm.Range(p(1).Offset(nrow + 3, 0), shtLm.UsedRange.SpecialCells(xlCellTypeLastCell).Offset(10, 0)).Clear
  
  If ycount > 0 Then
    Set p(3) = p(1).Offset(nrow + 4, 0)
    With p(3).Resize(ycount, ycount)
      .Value = 0
      .name = "LM!Ty"
      .Style = "Input"
      .Font.Bold = True
      .Font.Color = RGB(0, 0, 0)
    End With
    p(3).Offset(-1, 0).Resize(1, ycount).FormulaArray = "=" & p(1).Offset(1, 0).Resize(1, ycount).AddressLocal
    For i = 0 To ycount - 1
      p(3).Offset(i, i).Value = 1
    Next i
  End If

  If xcount > 0 Then
    Set p(4) = p(1).Offset(nrow + 4, ycount + ccount)
    With p(4).Resize(xcount, xcount)
      .Value = 0
      .name = "LM!Tx"
      .Style = "Input"
      .Font.Bold = True
      .Font.Color = RGB(0, 0, 0)
    End With
    p(4).Offset(-1, 0).Resize(1, xcount).FormulaArray = "=" & p(1).Offset(1, ycount + ccount).Resize(1, xcount).AddressLocal
    For i = 0 To xcount - 1
      p(4).Offset(i, i).Value = 1
    Next i
  End If
  
  shtLm.Range("Ty_row").Value = ycount
  shtLm.Range("Tx_row").Value = xcount
  
  shtLm.Range(p(1).Offset(0, ncol + 1), shtLm.UsedRange.SpecialCells(xlCellTypeLastCell).Offset(10, 0)).Clear
  t.HeaderRowRange.EntireColumn.HorizontalAlignment = xlCenter
lbl_exit:
  tmp.Clear
  
  
  Application.ScreenUpdating = states(1)
  Application.EnableEvents = states(2)
End Sub

Sub gen_WorkData()
'Require: data paste from pivot has been speciefied y-c-x
'Require: Range("Ty") and Range("Tx") been specified
  Dim n As Integer: n = shtLm.Range("n").Value
  Dim ycount As Integer: ycount = shtLm.Range("p").Value
  Dim ccount As Integer: ccount = shtLm.Range("kc").Value
  Dim xcount As Integer: xcount = shtLm.Range("k").Value
  Dim Tycount As Integer: Tycount = shtLm.Range("Ty_row").Value
  Dim xcountSelected As Integer: xcountSelected = shtLm.Range("Tx_row").Value
  Dim t As ListObject: Set t = shtLm.ListObjects("DataArea_Lm")
  Dim p(1 To 10) As Range, R As Range, r2 As Range, r3 As Range, c(1 To 100) As Range, h As Range
  Set p(1) = shtLm.Range("Left2Data").Offset(0, t.HeaderRowRange.Columns.count + 2)
  Dim i As Integer, j As Integer, k As Integer
  shtLm.Range(shtLm.Range("Left2Data").Offset(-1, t.HeaderRowRange.Columns.count + 1), shtLm.UsedRange.SpecialCells(xlCellTypeLastCell).Offset(10, 0)).Clear
  p(1).Value = "Working Data"
  p(1).HorizontalAlignment = xlRight
  Set R = p(1)
  
  Set p(2) = R.Offset(0, 1)
  For i = 1 To Tycount
    Set R = R.Offset(0, 1)
    'r.Value = "Ty" & i
    R.FormulaLocal = "=LinearCombinationOfStrings(" & t.Range.Cells(1, 1).Resize(1, ycount).AddressLocal & ", " & shtLm.Range("Ty").Rows(i).AddressLocal & ")"
    R.AddComment Text:=R.Value & " : " & "Ty" & i
  Next i
  shtLm.Range(p(2), R).Offset(1, 0).Resize(n, Tycount).FormulaArray = "=MMULT(" & shtLm.Range("Left2Data").Offset(1, 1).Resize(n, ycount).AddressLocal & ",TRANSPOSE(Ty))"
  
  Set p(2) = R.Offset(0, 1)
  For i = 1 To ccount
    Set R = R.Offset(0, 1)
    'r.Value = "c" & i
    R.FormulaLocal = "=" & t.Range.Cells(1, ycount + i).AddressLocal
    R.AddComment Text:=R.Value & " : " & "c" & i
    Set c(i) = R
  Next i
  If ccount > 0 Then
    shtLm.Range(p(2), R).Offset(1, 0).Resize(n, ccount).FormulaArray = "=" & shtLm.Range("Left2Data").Offset(1, ycount + 1).Resize(n, ccount).AddressLocal
  End If
  
  R.Offset(n + 2, 0).Value = "Intercept"
  For i = 1 To Tycount
    With R.Offset(n + 2 + i * 3 - 2, -1)
      .Value = "Beta for Ty" & i
      .HorizontalAlignment = xlRight
    End With
    With R.Offset(n + 2 + i * 3 - 1, -1)
      .Value = "sd"
      .HorizontalAlignment = xlRight
    End With
  Next i
  Set p(2) = R.Offset(0, 1)
  If xcountSelected > 0 Then
    For i = 1 To xcountSelected
      Set R = R.Offset(0, 1)
      'r.Value = "Tx" & i
      R.FormulaLocal = "=LinearCombinationOfStrings(" & t.Range.Cells(1, ycount + ccount + 1).Resize(1, xcount).AddressLocal & ", " & shtLm.Range("Tx").Rows(i).AddressLocal & ")"
      R.AddComment Text:=R.Value & " : " & "Tx" & i
      R.Offset(n + 2, 0).FormulaLocal = "=" & R.AddressLocal
      R.Offset(n + 2, 0).AddComment Text:=R.Comment.Text
    Next i
    shtLm.Range(p(2), R).Offset(1, 0).Resize(n, xcountSelected).FormulaArray = "=MMULT(" & shtLm.Range("Left2Data").Offset(1, ycount + ccount + 1).Resize(n, xcount).AddressLocal & ",TRANSPOSE(Tx))"
  End If
  Dim valuestr As String
  If ccount > 0 Then
'deal with categorical variables c(1) ~ c(ccount), assume ccount <= 100
    
    For k = 1 To ccount
      Set R = R.Offset(0, 1)
      Set h = t.HeaderRowRange.Cells(1, ycount + k)
      t.Sort.SortFields.Clear
      t.Sort.SortFields.Add Key:=t.Range.Columns(ycount + k), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
      With t.Sort
          .Header = xlYes
          .MatchCase = False
          .Orientation = xlTopToBottom
          .SortMethod = xlPinYin
          .Apply
      End With
      Set r2 = c(k).Offset(1, 0)
      j = 1
      If r2.Value <> c(k).Offset(n, 0).Value Then
        Set r3 = R.Offset(1, 0)
      End If
      
      'r.Value = "c" & k & j
      R.FormulaLocal = "=" & c(k).AddressLocal & " & " & j
      R.AddComment Text:=c(k).Value & " == " & r2.Value
      R.Offset(n + 2).FormulaLocal = "=" & R.AddressLocal
      R.Offset(n + 2).AddComment Text:=R.Comment.Text
      
      valuestr = CStr(r2.Offset(1, 0).Value)
      If TypeName(r2.Offset(1, 0).Value) <> "Double" Then
        valuestr = """" & valuestr & """"
      End If
      shtLm.Range(R.Offset(1, 0), R.Offset(n, 0)).FormulaLocal = "=IF(DataArea_Lm[@" & h.Value & "]=" & valuestr & ",1,0)"
      While r2.Offset(1, 0).Value <> c(k).Offset(n, 0).Value
        If r2.Offset(1, 0).Value <> r2.Value Then
          j = j + 1
          Set R = R.Offset(0, 1)
          'r.Value = "c" & k & j
          R.FormulaLocal = "=" & c(k).AddressLocal & " & " & j
          R.AddComment Text:=c(k).Value & " == " & r2.Offset(1, 0).Value
          R.Offset(n + 2).AddComment Text:=R.Comment.Text
          
          R.Offset(n + 2).FormulaLocal = "=" & R.AddressLocal
          
          valuestr = CStr(r2.Offset(1, 0).Value)
          If TypeName(r2.Offset(1, 0).Value) <> "Double" Then
            valuestr = """" & valuestr & """"
          End If
          shtLm.Range(R.Offset(1, 0), R.Offset(n, 0)).FormulaLocal = "=IF(DataArea_Lm[@" & h.Value & "]=" & valuestr & ",1,0)"
        End If
        Set r2 = r2.Offset(1, 0)
      Wend
        
    Next k
  End If
  
  
  Set p(3) = p(1).End(xlToRight).Offset(n, 0)
  shtLm.Range(p(1).Offset(0, 1), p(3)).name = "'" & shtLm.name & "'!WorkingData"
  shtLm.Range(p(1).Offset(1, 1), p(3)).name = "'" & shtLm.name & "'!WorkingDataBody"
  shtLm.Range(p(1).Offset(-1, 1), p(3)).HorizontalAlignment = xlCenter
  shtLm.Range(p(1).Offset(1, 1), p(1).Offset(n, Tycount)).name = "'" & shtLm.name & "'!TyData"
  shtLm.Range(p(1).Offset(1, Tycount + ccount + 1), p(3)).name = "'" & shtLm.name & "'!TxData"
  'beta_start
  p(1).Offset(n + 2, Tycount + ccount - 1).name = "'" & shtLm.name & "'!beta_start"
  shtLm.Range("beta_start").Offset(0, 1).End(xlToRight).Offset(0, 2).Value = "Rsquare"
  shtLm.Range("beta_start").Offset(0, 1).End(xlToRight).Offset(0, 2).name = "'" & shtLm.name & "'!RegRsquare"
  
  shtLm.Range("beta_start").Offset(0, 1).End(xlToRight).Offset(0, 3).Value = "F(anova)"
  shtLm.Range("beta_start").Offset(0, 1).End(xlToRight).Offset(0, 3).name = "'" & shtLm.name & "'!RegF"
  
  shtLm.Range("beta_start").Offset(0, 1).End(xlToRight).Offset(0, 4).Value = "F: df1"
  shtLm.Range("beta_start").Offset(0, 1).End(xlToRight).Offset(0, 4).name = "'" & shtLm.name & "'!RegF_df1"
  
  shtLm.Range("beta_start").Offset(0, 1).End(xlToRight).Offset(0, 5).Value = "F: df2"
  shtLm.Range("beta_start").Offset(0, 1).End(xlToRight).Offset(0, 5).name = "'" & shtLm.name & "'!RegF_df2"
  
  shtLm.Range("beta_start").Offset(0, 1).End(xlToRight).Offset(0, 6).Value = "F: pvalue"
  shtLm.Range("beta_start").Offset(0, 1).End(xlToRight).Offset(0, 6).name = "'" & shtLm.name & "'!RegF_pvalue"
  

  Set R = shtLm.Range("WorkingData").Cells(1, 1).End(xlToRight)
  Set p(4) = shtLm.Range("beta_start").Offset(0, 2)
  Set p(4) = shtLm.Range(p(4), p(4).Offset(0, -1).End(xlToRight))
  'Ty1_Pred
  R.Offset(1, 1).Resize(n, Tycount).name = "Ty_Pred"
  R.Offset(-1, 1).Resize(1, Tycount).Value = "Predicted"
  For j = 1 To Tycount
    Set R = R.Offset(0, 1)
    'r.Value = "Ty" & j & "_Pred"
    R.FormulaLocal = "=" & shtLm.Range("TyData").Columns(j).Cells(0, 1).AddressLocal
    '=MMULT(TxData,TRANSPOSE(M43:P43))+L43
    R.Offset(1, 0).Resize(n, 1).FormulaArray = "=MMULT(TxData,TRANSPOSE(" & p(4).Offset(1 + (j - 1) * 3, 0).AddressLocal & "))+" & shtLm.Range("beta_start").Offset(1 + (j - 1) * 3, 1).AddressLocal
  Next j
  
  'Ty1_Resid
  R.Offset(1, 1).Resize(n, Tycount).name = "Ty_Resid"
  R.Offset(-1, 1).Resize(1, Tycount).Value = "Residual"
  For j = 1 To Tycount
    Set R = R.Offset(0, 1)
    'r.Value = "Ty" & j & "_Resid"
    R.FormulaLocal = "=" & shtLm.Range("TyData").Columns(j).Cells(0, 1).AddressLocal
    R.Offset(1, 0).Resize(n, 1).FormulaArray = "=" & shtLm.Range("TyData").Columns(j).AddressLocal & "-" & R.Offset(1, -Tycount).Resize(n, 1).AddressLocal
  Next j
  t.HeaderRowRange.EntireRow.VerticalAlignment = xlTop
  shtLm.Range(shtLm.Range("WorkingData").Cells(1, 1), shtLm.Range("WorkingData").Cells(1, 1).End(xlToRight)).WrapText = True
  shtLm.Range("TxData").Rows(1).Offset(-2, 0).Style = "Input"
  With shtLm.Range("TxData").Rows(1).Offset(-2, 0).Cells(1, 0)
    .Value = "Put an ""x"" to select as a regressor->"
    .Font.Bold = True
    .HorizontalAlignment = xlRight
  End With
End Sub

Private Sub reg()
  Dim n As Long: n = shtLm.Range("n").Value
  Dim Tycount As Integer: Tycount = shtLm.Range("Ty_row").Value
  Dim xcountSelected As Integer: xcountSelected = 0
  Dim h As Range: Set h = shtLm.Range("TxData").Rows(1).Offset(-2, 0)
  Dim i As Long, j As Long
  Dim y, x, res
  Dim R As Range
  Set R = shtLm.Range("beta_start").Offset(1, 1)
  shtLm.Range(R, R.Offset(Tycount * 3 + 1, xcountSelected + 100)).Clear
  Set R = R.Offset(Tycount * 3, shtLm.Range("WorkingData").Cells(1, 1).Column - R.Column - 1)
  shtLm.Range(R, R.Offset(Tycount * 5 + 10, xcountSelected + 100)).Clear
  
  Dim xid(1 To 200) As Integer
  For Each R In h.Cells
    If LCase(Trim(R.Value)) = "x" Then
      xcountSelected = xcountSelected + 1
      xid(xcountSelected) = R.Column - h.Cells(1, 1).Column + 1
    End If
  Next R
  
  If xcountSelected = 0 Then GoTo lbl_exit
  
  ReDim x(1 To n, 1 To xcountSelected)
  For i = 1 To n
    For j = 1 To xcountSelected
      x(i, j) = shtLm.Range("TxData").Cells(i, xid(j)).Value
    Next j
  Next i
  
  For j = 1 To Tycount
    y = shtLm.Range("TyData").Columns(j).Value2
    res = MultipleRegression(y, x, True)
    Set R = shtLm.Range("beta_start").Offset(1 + (j - 1) * 3, 1)
    With R.Offset(0, 1).Resize(1, shtLm.Range("TxData").Columns.count)
      .Value = 0
      .Font.Color = RGB(255, 255, 255)
    End With
    
    R.Value = res(1, 1)
    With R
      .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=NORMSINV(1-alpha/2) * " & .Offset(1, 0).AddressLocal, Formula2:="=-NORMSINV(1-alpha/2) * " & .Offset(1, 0).AddressLocal
      .FormatConditions(.FormatConditions.count).SetFirstPriority
      With .FormatConditions(1).Font
          .Color = -16752384
          .TintAndShade = 0
      End With
      With .FormatConditions(1).Interior
          .PatternColorIndex = xlAutomatic
          .Color = 13561798
          .TintAndShade = 0
      End With
      .FormatConditions(1).StopIfTrue = False
    End With
    
    R.Offset(1, 0).Value = res(1, xcountSelected + 3)
    For i = 1 To xcountSelected
      With R.Offset(0, xid(i))
        .Value = res(1, i + 1)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=NORMSINV(1-alpha/2) * " & .Offset(1, 0).AddressLocal, Formula2:="=-NORMSINV(1-alpha/2) * " & .Offset(1, 0).AddressLocal
        .FormatConditions(.FormatConditions.count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = -16752384
            .TintAndShade = 0
        End With
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13561798
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
        
        
        .Font.Color = RGB(0, 0, 0)
      End With
      R.Offset(1, xid(i)).Value = res(1, i + xcountSelected + 3)
    Next i
    Set R = shtLm.Range("beta_start").Offset(0, 1).End(xlToRight).End(xlToRight).Offset(1 + (j - 1) * 3, 0)
    R.Value = res(1, xcountSelected + 2)
    R.Offset(0, 1).Value = res(1, (1 + xcountSelected) * 2 + 2)
    R.Offset(0, 2).Value = res(1, (1 + xcountSelected) * 2 + 3)
    R.Offset(0, 3).Value = res(1, (1 + xcountSelected) * 2 + 4)
    R.Offset(0, 4).FormulaLocal = "=FDIST(" & R.Offset(0, 1).AddressLocal & ", " & R.Offset(0, 2).AddressLocal & ", " & R.Offset(0, 3).AddressLocal & ")"
    Call pvalue_Formatting(R.Offset(0, 4), "alpha")
  Next j
  shtLm.Range(shtLm.Range("beta_start").Offset(3 * Tycount + 1, 1), shtLm.Range("beta_start").Offset(0, 1).End(xlToRight).Offset(0, 10)).HorizontalAlignment = xlCenter
  
  
  
  
'==== manova, using temp sheet to compute the E-CSSP and H-CSSP ======
  Dim m(1 To 10) As Range
  Set m(1) = shtLm.Range("WorkingData").Cells(1, 1).Offset(n + 3 * Tycount + 3, -1)
  m(1).Value = "MANOVA"
  m(1).Font.Bold = True
  shtLm.Range("manova_g").Value = xcountSelected + 1
  shtLm.Range("manova_a").FormulaLocal = "=n-manova_g"
  shtLm.Range("manova_b").FormulaLocal = "=manova_g-1"
'  Dim cssp_mat, cIndex, datamat
'
'  ReDim cIndex(1 To shtLm.Range("WorkingDataBody").Columns.count)
'  i = 0
'  For Each r In h.Cells
'    If LCase(Trim(r.Value)) = "x" Then
'      i = i + 1
'      cIndex(i) = r.Column - h.Cells(1, 1).Column + 1 + Tycount
'    End If
'  Next r
'  ReDim Preserve cIndex(1 To i)
'  shtLm.Range("TyData,TxData").Copy
'  Set r = shtTemp.UsedRange(shtTemp.UsedRange.count).Offset(2, 1).End(xlToLeft)
'  r.PasteSpecial Paste:=xlValues
'
'  Application.CutCopyMode = False
'  cssp_mat = CSSP(r.Resize(n, Tycount + h.Columns.count).Value2, cIndex, r)
'
'  shtLm.Range("manova_g").Value = cssp_mat(4)
  
  
  
  Set m(2) = m(1).Offset(1, 0)
  m(2).Value = "Ty Overall Mean"
  m(2).Font.Bold = True
  m(2).Resize(Tycount * 3, 1).HorizontalAlignment = xlRight
  For j = 1 To Tycount
    m(2).Offset(0, j).FormulaLocal = "=AVERAGE(" & shtLm.Range("TyData").Columns(j).AddressLocal & ")"
  Next j
  m(2).Offset(0, 1).Resize(1, Tycount).name = "Ty_Overall_Mean"
  
  Set m(2) = m(2).Offset(2, 0)
  m(2).Value = "E - CSSP"
  m(2).Font.Bold = True
  
  Set m(3) = m(2).Offset(0, 1).Resize(Tycount, Tycount)
  m(3).FormulaArray = "=MMULT(TRANSPOSE(Ty_Resid),Ty_Resid)"
  m(3).name = "'" & shtLm.name & "'!E_CSSP"
  
  Set m(4) = m(2).Offset(Tycount + 1, 0)
  m(4).Value = "H - CSSP"
  m(4).Font.Bold = True
  
  Set m(5) = m(4).Offset(0, 1).Resize(Tycount, Tycount)
  m(5).FormulaArray = "=MMULT(TRANSPOSE(Ty_Pred),Ty_Pred) - n * MMULT(TRANSPOSE(Ty_Overall_Mean),Ty_Overall_Mean)"
  m(5).name = "'" & shtLm.name & "'!H_CSSP"
  
  Set m(6) = m(1).Offset(0, Tycount + 2)
  m(6).Value = "Wilks' Lambda"
  m(6).Font.Bold = True
  
  m(6).Offset(1, 0).Value = "df1"
  m(6).Offset(1, 1).FormulaLocal = "=Ty_row*manova_b"
  
  m(6).Offset(2, 0).Value = "df2"
  m(6).Offset(2, 1).FormulaLocal = "=manova_r*manova_t-manova_q"
  
  m(6).Offset(3, 0).Value = "stat"
  m(6).Offset(3, 1).FormulaLocal = "=MDETERM(E_CSSP)/MDETERM(E_CSSP+H_CSSP)"
  m(6).Offset(3, 1).Style = "Note"
  
  
  m(6).Offset(4, 0).Value = "Fval"
  m(6).Offset(4, 1).FormulaLocal = "=" & m(6).Offset(2, 1).AddressLocal & "/" & m(6).Offset(1, 1).AddressLocal & _
                                   "*(1-" & m(6).Offset(3, 1).AddressLocal & "^(1/manova_t))/" & m(6).Offset(3, 1).AddressLocal & "^(1/manova_t)"
    
  m(6).Offset(5, 0).Value = "pvalue"
  m(6).Offset(5, 1).FormulaLocal = "=FDIST(" & m(6).Offset(4, 1).AddressLocal & "," & _
                                    m(6).Offset(1, 1).AddressLocal & "," & _
                                    m(6).Offset(2, 1).AddressLocal & ")"
  Call pvalue_Formatting(m(6).Offset(5, 1), "alpha")
  
  
  Set m(6) = m(6).Offset(0, 3)
  m(6).Value = "Bartlett-Nanda-Pillai Trace"
  m(6).Font.Bold = True
  
  m(6).Offset(1, 0).Value = "df1"
  m(6).Offset(1, 1).FormulaLocal = "=manova_s*(2*manova_m1+manova_s+1)"
  
  m(6).Offset(2, 0).Value = "df2"
  m(6).Offset(2, 1).FormulaLocal = "=manova_s*(2*manova_m2+manova_s+1)"
  
  m(6).Offset(3, 0).Value = "stat"
  m(6).Offset(3, 1).FormulaLocal = "=MTr(MMULT(MINVERSE(E_CSSP+H_CSSP),H_CSSP))"
  m(6).Offset(3, 1).Style = "Note"
  
  m(6).Offset(4, 0).Value = "Fval"
  m(6).Offset(4, 1).FormulaLocal = "=" & m(6).Offset(2, 1).AddressLocal & "/" & m(6).Offset(1, 1).AddressLocal & _
                                   "*" & m(6).Offset(3, 1).AddressLocal & "/(manova_s-" & m(6).Offset(3, 1).AddressLocal & ")"
    
  m(6).Offset(5, 0).Value = "pvalue"
  m(6).Offset(5, 1).FormulaLocal = "=FDIST(" & m(6).Offset(4, 1).AddressLocal & "," & _
                                    m(6).Offset(1, 1).AddressLocal & "," & _
                                    m(6).Offset(2, 1).AddressLocal & ")"
  Call pvalue_Formatting(m(6).Offset(5, 1), "alpha")
  
  
  Set m(6) = m(6).Offset(0, 3)
  m(6).Value = "Hotelling-Lawley Trace"
  m(6).Font.Bold = True
  
  m(6).Offset(1, 0).Value = "df1"
  m(6).Offset(1, 1).FormulaLocal = "=IF(manova_m2>0,Ty_row*manova_b,manova_s*(2*manova_m1+manova_s+1))"
  
  m(6).Offset(2, 0).Value = "df2"
  m(6).Offset(2, 1).FormulaLocal = "=IF(manova_m2>0,4+(Ty_row*manova_b+2)/(manova_d-1),2*(manova_s*manova_m2+1))"
  
  m(6).Offset(3, 0).Value = "stat"
  m(6).Offset(3, 1).FormulaLocal = "=MTr(MMULT(MINVERSE(E_CSSP),H_CSSP))"
  m(6).Offset(3, 1).Style = "Note"
  
  m(6).Offset(4, 0).Value = "Fval"
  '=IF(manova_m2>0,S33/S32*S34,S33/S32*S34/manova_s)
  m(6).Offset(4, 1).FormulaLocal = "=IF(manova_m2>0," & _
                m(6).Offset(2, 1).AddressLocal & "/" & m(6).Offset(1, 1).AddressLocal & "*" & m(6).Offset(3, 1).AddressLocal & "," & _
                m(6).Offset(2, 1).AddressLocal & "/" & m(6).Offset(1, 1).AddressLocal & "*" & m(6).Offset(3, 1).AddressLocal & "/manova_s)"
  m(6).Offset(5, 0).Value = "pvalue"
  m(6).Offset(5, 1).FormulaLocal = "=FDIST(" & m(6).Offset(4, 1).AddressLocal & "," & _
                                    m(6).Offset(1, 1).AddressLocal & "," & _
                                    m(6).Offset(2, 1).AddressLocal & ")"
  Call pvalue_Formatting(m(6).Offset(5, 1), "alpha")
  
  
  
  Set m(6) = m(6).Offset(0, 3)
  m(6).Value = "Roy's Greatest Root"
  m(6).Font.Bold = True
  
  m(6).Offset(1, 0).Value = "df1"
  m(6).Offset(1, 1).FormulaLocal = "=manova_h"
  
  m(6).Offset(2, 0).Value = "df2"
  m(6).Offset(2, 1).FormulaLocal = "=manova_a+manova_b-manova_h"
  
  m(6).Offset(3, 0).Value = "stat"
  m(6).Offset(3, 1).FormulaLocal = "=CovEigenvalueMax(MMULT(MINVERSE(E_CSSP),H_CSSP))"
  m(6).Offset(3, 1).Style = "Note"
  
  m(6).Offset(4, 0).Value = "Fval"
  '=IF(manova_m2>0,S33/S32*S34,S33/S32*S34/manova_s)
  m(6).Offset(4, 1).FormulaLocal = "=" & _
        m(6).Offset(2, 1).AddressLocal & "/" & m(6).Offset(1, 1).AddressLocal & "*" & m(6).Offset(3, 1).AddressLocal
  m(6).Offset(5, 0).Value = "pvalue"
  m(6).Offset(5, 1).FormulaLocal = "=FDIST(" & m(6).Offset(4, 1).AddressLocal & "," & _
                                    m(6).Offset(1, 1).AddressLocal & "," & _
                                    m(6).Offset(2, 1).AddressLocal & ")"
  Call pvalue_Formatting(m(6).Offset(5, 1), "alpha")
  
lbl_exit:
End Sub
