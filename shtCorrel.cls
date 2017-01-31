VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtCorrel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  On Error GoTo lbl_exit
  
  If sameRange(Target, shtCorrel.Range("simple_Data")) Then
    Cancel = True
    
    Call simple_Data
    
    GoTo lbl_exit
  End If
  
  
  If sameRange(Target, shtCorrel.Range("simple_DeterminePar")) Then
    Cancel = True
    
    Call simple_DeterminePar
    
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtCorrel.Range("paste_Data_Matrix")) Then
    Cancel = True
    
    Call paste_Data_Matrix
    
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtCorrel.Range("clear_all")) Then
    Cancel = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Call clear_correl
    
    GoTo lbl_exit
  End If
 
  If sameRange(Target, shtCorrel.Range("clear_multiple_regression_data")) Then
    Cancel = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Call clear_Data_Region
    
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtCorrel.Range("bi")) Then
    Cancel = True
    
    Call compute_multiple_regression
    
    GoTo lbl_exit
  End If
lbl_exit:
  Application.EnableEvents = True
  Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  On Error GoTo lbl_exit
  Dim states(1 To 2) As Boolean
  states(1) = Application.EnableEvents:  Application.EnableEvents = False
  states(2) = Application.ScreenUpdating: Application.ScreenUpdating = False
  
  If Target.Row = shtCorrel.Range("TopLeft_to_Datamatrix").Row And _
     Target.Column >= shtCorrel.Range("DataArea_Correl[#Headers]").Cells(1, 1).Column And _
     Target.Column < shtCorrel.Range("DataArea_Correl[#Headers]").Cells(1, 1).Column + shtCorrel.Range("multiple_p").Value _
  Then
      Call compute_multiple_regression
      GoTo lbl_exit
  End If
lbl_exit:
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub


Private Sub clear_correl()
  On Error GoTo lbl_exit
  Dim states(1 To 2) As Boolean
  states(1) = Application.EnableEvents:  Application.EnableEvents = False
  states(2) = Application.ScreenUpdating: Application.ScreenUpdating = False
  Dim nameArray: nameArray = Array("simple_rho0", "partial_rho0", "partial_n", "partial_q", "partial_r")
  Dim x
  For Each x In nameArray
    shtCorrel.Range(CStr(x)).ClearContents
  Next x
  Dim R As Range, last As Range
  Set R = shtCorrel.Range("simple_x")
  Set last = shtCorrel.UsedRange(shtCorrel.UsedRange.count)
  Set last = Intersect(R.EntireRow, last.EntireColumn)
  shtCorrel.Range(R.Offset(1, 1), last).Clear

  Call clear_Data_Region
  R.Select
  Set last = shtCorrel.UsedRange(shtCorrel.UsedRange.count)
  Set R = shtCorrel.Range("A1000").End(xlUp).Offset(1, 0)
  If last.Row >= R.Row Then
    shtCorrel.Range(R, last).EntireRow.Delete
  End If
  
  Set last = shtCorrel.UsedRange(shtCorrel.UsedRange.count)
  Dim p As Range
  Set R = [AAA1].End(xlToLeft).Offset(0, 1)
  For Each p In shtCorrel.Range(last.End(xlUp), last)
    If p.End(xlToLeft).Column > R.Column Then
      Set R = p.End(xlToLeft).Offset(0, 1)
    End If
  Next p
  If last.Column >= R.Column Then
    shtCorrel.Range(R, last).EntireColumn.Delete
  End If
  
  shtCorrel.UsedRange.ColumnWidth = 8.43
lbl_exit:
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub

Private Sub clear_Data_Region()
On Error GoTo lbl_exit
  Dim states(1 To 2) As Boolean
  states(1) = Application.EnableEvents:  Application.EnableEvents = False
  states(2) = Application.ScreenUpdating: Application.ScreenUpdating = False
  Dim x
  For Each x In shtCorrel.ListObjects
    x.Delete
  Next x
  For Each x In shtCorrel.ChartObjects
    x.Delete
  Next x
  
  Dim R As Range: Set R = shtCorrel.Range("sd_of_bi").Offset(0, 1)
  Dim last As Range: Set last = shtCorrel.UsedRange(shtCorrel.UsedRange.count)
  If last.Column >= R.Column And last.Row >= R.Row Then
    shtCorrel.Range(R, last).Clear
    shtCorrel.Range(R, last).ClearOutline
  Else
    R.Resize(100, 100).Clear
    R.Resize(100, 100).ClearOutline
  End If
  
  For Each x In Array("multiple_n", "multiple_p", "multiple_r", "mutiple_rsq", "multiple_F", "multiple_SSR", "multiple_df_SSR", "multiple_SSE", "multiple_df_SSE")
    shtCorrel.Range(CStr(x)).Value = ""
  Next x
  With shtCorrel.Cells
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
  End With
lbl_exit:
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub

Private Sub compute_multiple_regression()
  On Error GoTo lbl_exit
  Dim states(1 To 2) As Boolean
  states(1) = Application.EnableEvents:  Application.EnableEvents = False
  states(2) = Application.ScreenUpdating: Application.ScreenUpdating = False
  Dim originalSelection As Range
  Set originalSelection = Selection
  Dim i As Long, j As Long, k As Long
  Dim p As Integer:   p = shtCorrel.Range("DataArea_Correl[#Data]").Columns.count
  'shtCorrel.Range("multiple_p").Value = p
  Dim bi As Range: Set bi = shtCorrel.Range("bi")
  
  Dim sdbi As Range: Set sdbi = shtCorrel.Range("sd_of_bi")
  Dim h As Range: Set h = shtCorrel.Range("TopLeft_to_Datamatrix").Offset(1, 1)
  Dim num_of_y As Integer:   num_of_y = 0
  Dim dataRange As Range:
  Set dataRange = shtCorrel.UsedRange(shtCorrel.UsedRange.count).Offset(3, 0)
  Set dataRange = shtCorrel.Range("DataArea_Correl[#Data]").SpecialCells(xlCellTypeVisible)
  dataRange.Copy
  shtCorrel.Range("TopLeft_to_Datamatrix").Offset(shtCorrel.Range("DataArea_Correl[#Data]").Rows.count + 100, 1).PasteSpecial Paste:=xlValues
  Set dataRange = Selection
  Application.CutCopyMode = False
  Dim n As Integer:   n = dataRange.Rows.count
  shtCorrel.Range("multiple_n").Value = n
  shtCorrel.Range(bi.Offset(0, 1), bi.Offset(0, p)).Value = ""
  shtCorrel.Range(sdbi.Offset(0, 1), sdbi.Offset(0, p)).Value = ""
  shtCorrel.Range("TopLeft_to_Datamatrix").Offset(0, p + 2).Resize(p * 5, p * 2).ClearContents
  Dim y: ReDim y(1 To n, 1 To 1)
  Dim x: ReDim x(1 To n, 1 To p)
  
  Dim xcount As Integer:    xcount = 0
  Dim b: ReDim b(0 To p) As Integer
  
  k = 0
  For Each h In shtCorrel.Range("DataArea_Correl[#Headers]").Cells
    k = h.Column - shtCorrel.Range("TopLeft_to_Datamatrix").Column
    If LCase(Trim(h.Offset(-1, 0).Value)) = "y" Then
      num_of_y = num_of_y + 1
      If num_of_y >= 2 Then
        'MsgBox "You have specified more than one ys."
        h.Offset(-1, 0).Select
        Exit Sub
      End If
      y = dataRange.Columns(k).Value2
      
      b(0) = k
    ElseIf LCase(Trim(h.Offset(-1, 0).Value)) = "x" Then
      xcount = xcount + 1
      b(xcount) = k
      For i = 1 To n
        x(i, xcount) = dataRange.Cells(i, k).Value
      Next i
    End If
  
  Next h
  ReDim Preserve x(1 To n, 1 To xcount)
  Dim res: res = Application.LinEst(y, x, True, True)
  
  Dim Coef: ReDim Coef(0 To xcount, 1 To 1) As Double
  For i = 0 To xcount
    bi.Offset(0, b(i)).Value = res(1, 1 + xcount - i)
    Coef(i, 1) = res(1, 1 + xcount - i)
    sdbi.Offset(0, b(i)).Value = res(2, 1 + xcount - i)
  Next i
  shtCorrel.Range("mutiple_rsq").Value = res(3, 1)
  shtCorrel.Range("multiple_F").Value = res(4, 1)
  shtCorrel.Range("multiple_df_SSR").Value = xcount
  shtCorrel.Range("multiple_df_SSE").Value = res(4, 2)
  shtCorrel.Range("multiple_SSR").Value = res(5, 1)
  shtCorrel.Range("multiple_SSE").Value = res(5, 2)
  Dim x_aug: ReDim x_aug(1 To n, 0 To xcount) As Double
  
  
  For i = 1 To n
    x_aug(i, 0) = 1
    For j = 1 To xcount
      x_aug(i, j) = x(i, j)
    Next j
  Next i
  shtCorrel.Range("DataArea_Correl[y_predicted]").ClearContents
  shtCorrel.Range("DataArea_Correl[Residual]").ClearContents
  Set h = shtCorrel.Range("DataArea_Correl[y_predicted]").SpecialCells(xlCellTypeVisible).Cells
  Dim temp: temp = Application.MMult(x_aug, Coef)
  Dim R As Range
  i = 0
  For Each R In h
    i = i + 1
    R.Value = temp(i, 1)
    'r.Offset(0, 1).Value = y(i, 1) - temp(i, 1)
  Next R
  
  'h.Value2 = Application.MMult(x_aug, Coef)
  shtCorrel.Range("DataArea_Correl[Residual]").FormulaR1C1 = "=[@" & shtCorrel.Range("TopLeft_to_Datamatrix").Offset(1, b(0)).Value & "] - [@[y_predicted]]"
  
  Set h = shtCorrel.Range("TopLeft_to_Datamatrix").Offset(shtCorrel.Range("DataArea_Correl[#Data]").Rows.count + 4, 1)
  shtCorrel.Range(h.Offset(250, 0).End(xlUp), h.Offset(0, 100).End(xlToLeft)).Clear
  h.Value = "Cov Mat"
  For i = 0 To xcount
    h.Offset(i + 1, 0).Value = shtCorrel.Range("TopLeft_to_Datamatrix").Offset(1, b(i)).Value
    h.Offset(0, i + 1).Value = shtCorrel.Range("TopLeft_to_Datamatrix").Offset(1, b(i)).Value
  Next i
  For i = 1 To n
    x_aug(i, 0) = y(i, 1)
  Next i
  
  Dim cov:   cov = MCOVARIANCE(x_aug)
  h.Offset(1, 1).Resize(xcount + 1, xcount + 1).Value2 = cov
  Set h = h.Offset(xcount + 4, 0)
  h.Value = "Correl Mat"
  For i = 0 To xcount
    h.Offset(i + 1, 0).Value = shtCorrel.Range("TopLeft_to_Datamatrix").Offset(1, b(i)).Value
    h.Offset(0, i + 1).Value = shtCorrel.Range("TopLeft_to_Datamatrix").Offset(1, b(i)).Value
  Next i
  Dim sdInv:   sdInv = Application.MInverse(MakeDiagMatrix(GetDiagVector(cov)))
  For i = 1 To xcount + 1
    sdInv(i, i) = Sqr(sdInv(i, i))
  Next i
  Dim corr: corr = Application.MMult(Application.MMult(sdInv, cov), sdInv)
  h.Offset(1, 1).Resize(xcount + 1, xcount + 1).Value2 = corr
  
  shtCorrel.Range("multiple_r").Formula = "=SQRT(MMULT(MMULT(" & h.Offset(1, 2).Resize(1, xcount).AddressLocal & ", MINVERSE(" & h.Offset(2, 2).Resize(xcount, xcount).AddressLocal & ")), " & h.Offset(2, 1).Resize(xcount, 1).AddressLocal & "))"
  
lbl_exit:
  dataRange.Clear
  originalSelection.Select
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub





Private Sub paste_Data_Matrix()
On Error GoTo lbl_exit
  Dim states(1 To 2) As Boolean
  states(1) = Application.EnableEvents:  Application.EnableEvents = False
  states(2) = Application.ScreenUpdating: Application.ScreenUpdating = False
  Call clear_Data_Region
  Dim R As Range: Set R = shtCorrel.Range("TopLeft_to_Datamatrix")
  
  Application.GoTo R, True
  
  Dim lst As ListObject: Set lst = shtPivot.ListObjects("DataArea")
  
  If lst.AutoFilter.FilterMode = False Then
    lst.Range.Copy
    R.Offset(1, 1).PasteSpecial Paste:=xlPasteAll
    R.Offset(1, 1).ListObject.name = "DataArea_Correl"
  Else
    lst.Range.SpecialCells(xlCellTypeVisible).Copy
    R.Offset(1, 1).PasteSpecial Paste:=xlPasteAll
    shtCorrel.ListObjects.Add(xlSrcRange, Selection, , xlYes).name = "DataArea_Correl"
  End If
  R.Offset(1, 1).ListObject.DataBodyRange.Rows.Group
  
  shtCorrel.Range("multiple_n").Value = shtCorrel.Range("DataArea_Correl[#Data]").Rows.count
  shtCorrel.Range("multiple_p").Value = shtCorrel.Range("DataArea_Correl[#Data]").Columns.count
  R.Offset(1, 1 + shtCorrel.Range("multiple_p").Value).Value = "y_predicted"
  R.Offset(1, 2 + shtCorrel.Range("multiple_p").Value).Value = "Residual"
lbl_exit:
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub

Private Sub simple_Data()
  On Error GoTo lbl_exit
  Dim states(1 To 2) As Boolean
  states(1) = Application.EnableEvents:  Application.EnableEvents = False
  states(2) = Application.ScreenUpdating: Application.ScreenUpdating = False
  
  Dim R As Range: Set R = shtCorrel.Cells.SpecialCells(xlCellTypeLastCell).End(xlToLeft).Offset(100, 0)
  Dim Q As Range
  R.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Set Q = Selection
  Dim t As Range:  Set t = Intersect(shtCorrel.Cells.SpecialCells(xlCellTypeLastCell).EntireColumn, shtCorrel.Range("simple_x").EntireRow)
  shtCorrel.Range(t, shtCorrel.Range("simple_y").Offset(0, 1)).ClearContents
  If R.Value <> "" And R.Offset(1, 0).Value <> "" And R.Offset(2, 0).Value = "" Then
    Selection.Copy
    shtCorrel.Range("simple_x").Offset(0, 1).PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  ElseIf R.Value <> "" And R.Offset(0, 1).Value <> "" And R.Offset(0, 2).Value = "" Then
    Selection.Copy
    shtCorrel.Range("simple_x").Offset(0, 1).PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
  Else
    GoTo lbl_exit:
  End If
  Call simple_DeterminePar
lbl_exit:
  R.Value = "."
  shtCorrel.Range(R.Offset(10000, 0).End(xlUp), R.Offset(0, 5000).End(xlToLeft)).Clear
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub

Private Sub simple_DeterminePar()
  Dim n As Long
  Dim x As Range: Set x = shtCorrel.Range("simple_x")
  n = x.End(xlToRight).Column - x.Column
  shtCorrel.Range("simple_n").Value = n
  Set x = shtCorrel.Range(shtCorrel.Range("simple_x").Offset(0, 1), shtCorrel.Range("simple_x").Offset(0, n))
  shtCorrel.Range("simple_r").Formula = "=CORREL(" & x.AddressLocal & "," & x.Offset(1, 0).AddressLocal & ")"
  shtCorrel.Range("Simple_Regression").FormulaArray = "=MultipleRegression(" & x.Offset(1, 0).AddressLocal & "," & x.AddressLocal & ")"
  
End Sub

'Private Sub toggle_data_view()
'  Dim r As Range: Set r = shtCorrel.Range("Multiple_Correl")
'  Set r = shtCorrel.Range(r.Offset(2, 0), shtCorrel.Range("TopLeft_to_Datamatrix").Offset(-1, 0)).EntireRow
'  If r.Hidden Then
'    r.Hidden = False
'  Else
'    r.Hidden = True
'  End If
'End Sub
'
'Private Sub viewSelector()
'  On Error GoTo lbl_exit:
'  'Application.Goto _
'          reference:=shtCovTest.Cells.Find(what:=shtCovTest.Range("selector").Value, after:=shtCovTest.Range("A1"), LookIn:= _
'            xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext _
'            , MatchCase:=True), _
'          Scroll:=True
'  Application.Goto _
'          reference:=find2(shtCorrel.Range("selector").Value, shtCorrel.Range("A1")), _
'          Scroll:=True
'lbl_exit:
'End Sub


