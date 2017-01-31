VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  On Error Resume Next
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  
  
  If sameRange(Target, shtPC.Range("clear")) Then
    Cancel = True
      
    Call clear_PC
      
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtPC.Range("Paste_from_Pivot")) Then
    Cancel = True
      
    Call Paste_from_Pivot
      
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtPC.Range("Select_with_x")) Then
    Cancel = True
      
    Call toggle_select_all
      
    GoTo lbl_exit
  End If
  
lbl_exit:
  Application.EnableEvents = True
  Application.ScreenUpdating = True
End Sub




Sub Paste_from_Pivot()
  Call clear_PC
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  Dim R As Range:   Set R = shtPC.Range("Left2Data").Offset(0, 1)
  Dim lst As ListObject: Set lst = shtPivot.ListObjects("DataArea")
  
  If lst.AutoFilter.FilterMode = False Then
    lst.Range.Copy
    R.PasteSpecial Paste:=xlPasteAll
    R.ListObject.name = "DataArea_PC"
  Else
    lst.Range.SpecialCells(xlCellTypeVisible).Copy
    R.PasteSpecial Paste:=xlPasteAll
    shtPC.ListObjects.Add(xlSrcRange, Selection, , xlYes).name = "DataArea_PC"
  End If
  R.ListObject.DataBodyRange.Rows.Group
  
  Application.EnableEvents = True
  Application.ScreenUpdating = True
End Sub

Private Sub toggle_select_all()
  If existObject("DataArea_PC", "ListObject", shtPC) = False Then Exit Sub
  
  Dim Q(1 To 3) As Range, R As Range
  Set Q(1) = shtPC.Range("Select_with_x").Offset(0, 1)
  Set Q(2) = Q(1).Offset(0, shtPC.ListObjects("DataArea_PC").HeaderRowRange.Columns.count - 1)
  For Each R In shtPC.Range(Q(1), Q(2)).Cells
    If R.Value <> "" Then
      GoTo lbl_clear
    End If
  Next R
  shtPC.Range(Q(1), Q(2)).Value = "x"
  GoTo lbl_exit
lbl_clear:
  shtPC.Range(Q(1), Q(2)).ClearContents
  GoTo lbl_exit
lbl_exit:
  Call compute_PC
End Sub


Sub clear_PC()
  Application.EnableEvents = False
  Dim x: For Each x In shtPC.ListObjects
    x.Delete
  Next x:
  
  For Each x In shtPC.ChartObjects:  x.Delete: Next x
  For Each x In Array("n", "p", "m", "CI_i", "CI_lambda_i", "total_var")
    shtPC.Range(CStr(x)).ClearContents
  Next x
  
  Dim R As Range
  Set R = shtPC.Range("Left2Data").Offset(-1, 1)
  Dim last As Range: Set last = shtPC.UsedRange(shtPC.UsedRange.count)
  If last.Column >= R.Column And last.Row >= R.Row Then
    shtPC.Range(R, last).Clear
    shtPC.Range(R, last).ColumnWidth = 8.43
  Else
    R.Resize(100, 100).Clear
    R.Resize(100, 100).ColumnWidth = 8.43
  End If
  Set R = shtPC.Range("Left2Data").Offset(1000, 0).End(xlUp).Offset(1, 0)
  Set last = shtPC.UsedRange(shtPC.UsedRange.count)
  If last.Row >= R.Row Then
    shtPC.Range(R, last).EntireRow.Delete
  End If
  Set R = shtPC.Range("AAA1").End(xlToLeft).Offset(0, 5)
  Set last = shtPC.UsedRange(shtPC.UsedRange.count)
  If last.Column >= R.Column Then
    shtPC.Range(R, last).EntireColumn.Delete
  End If
  shtPC.Range("tests_based_on").Value = "Correl"
  shtPC.UsedRange.ClearOutline
  Application.EnableEvents = True
End Sub

Sub compute_PC()
  On Error Resume Next
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  
  Dim i As Long, j As Long, k As Long, Q(1 To 20) As Range
  Set Q(1) = shtPC.Range("Left2Data").Offset(0, 1)
  Dim n As Long: n = Q(1).ListObject.DataBodyRange.Rows.count
  shtPC.Range("n") = n
  Dim p As Long: p = Q(1).ListObject.HeaderRowRange.count
  
  Q(1).Offset(0, p + 1).Value = "."
  shtPC.Range(Q(1).Offset(-1, p), shtPC.UsedRange(shtPC.UsedRange.count)).Clear
  Dim headers: ReDim headers(1 To p) As String
  Dim pos: ReDim pos(1 To p) As Integer
  Dim xcount As Integer: xcount = 0
  
  Dim xUnion As Range
  Dim x
  For Each x In shtPC.ListObjects("DataArea_PC").HeaderRowRange.Cells
    If Trim(x.Offset(-1, 0).Value) = "x" Then
      xcount = xcount + 1
      If xcount = 1 Then
        Set xUnion = x.Offset(-1, 0)
      ElseIf xcount >= 2 Then
        Set xUnion = Union(xUnion, x.Offset(-1, 0))
      End If
      pos(xcount) = x.Column - Q(1).Column + 1
      headers(xcount) = x.Value
    End If
  Next x
  shtPC.Range("p").Value = xcount
  Dim mu: ReDim mu(1 To xcount) As Double
  Dim sd: ReDim sd(1 To xcount) As Double
  Dim Data: ReDim Data(1 To n, 1 To xcount) 'de-meaned
  Dim data_stu: ReDim data_stu(1 To n, 1 To xcount) 'de-meaned
  For j = 1 To xcount
    mu(j) = Application.Average(shtPC.ListObjects("DataArea_PC").DataBodyRange.Columns(pos(j)).Value2)
    sd(j) = Application.StDev(shtPC.ListObjects("DataArea_PC").DataBodyRange.Columns(pos(j)).Value2)
    For i = 1 To n
      Data(i, j) = shtPC.ListObjects("DataArea_PC").DataBodyRange.Cells(i, pos(j)).Value - mu(j)  'de-meaned
      data_stu(i, j) = Data(i, j) / sd(j)
    Next i
  Next j
  Dim alleigval As Range, taileigval As Range, headeigval As Range, m As Integer
'======== PC based on Covariance Matrix ===
  Set Q(2) = Q(1).Offset(0, p + 2)
  Q(2).Value = "Cov"
  For j = 1 To xcount
    Q(2).Offset(0, j).Value = headers(j)
    Q(2).Offset(j, 0).Value = headers(j)
  Next j
  
  Q(2).Offset(1, 1).Resize(xcount, xcount).Value2 = MCOVARIANCE(Data)
  
  Set Q(3) = Q(2).Offset(xcount + 3)
  Q(3).Value = "Eigval"
  Q(3).Offset(0, 1).Resize(2 + xcount, xcount).FormulaArray = "=CovEigenDecompQR(" & Q(2).Offset(1, 1).Resize(xcount, xcount).AddressLocal & ")"
  
  '==== statistical tests about the eigenvalues ======='
  If shtPC.Range("tests_based_on").Value = "Cov" Then
      m = shtPC.Range("m").Value
      If m < 0 Or m >= xcount Then
        m = 0
        shtPC.Range("m").Value = 0
      End If
      Set alleigval = Q(3).Offset(0, 1).Resize(1, xcount)
      Set taileigval = Q(3).Offset(0, 1 + m).Resize(1, xcount - m)
      If m > 0 Then Set headeigval = Q(3).Offset(0, 1).Resize(1, m)
      shtPC.Range("total_var").Formula = "=SUM(" & alleigval.AddressLocal & ")"
      shtPC.Range("g_lambda").Formula = "=SUM(" & taileigval.AddressLocal & ") / total_var"
    
      If m = 0 Then
        shtPC.Range("var_pct").Value = ""
        shtPC.Range("aux").Formula = "= (1 - g_lambda) ^ 2 * SUMSQ(" & taileigval.AddressLocal & ")"
      Else
        shtPC.Range("var_pct").Formula = "=" & Q(3).Offset(0, m).AddressLocal & " / total_var"
        shtPC.Range("aux").Formula = "= g_lambda ^ 2 * SUMSQ(" & headeigval.AddressLocal & _
                              ") + (1 - g_lambda) ^ 2 * SUMSQ(" & taileigval.AddressLocal & ")"
      End If
      shtPC.Range("equal_small_eigval_chisq").FormulaArray = "=(n-1) *  ((p-m)*LN(SUM(" & taileigval.AddressLocal & _
                                                                   ")/(p-m)) - SUM(LN(" & taileigval.AddressLocal & ")))"
      

      i = shtPC.Range("CI_i").Value
      If i >= 1 And i <= xcount Then
        shtPC.Range("CI_lambda_i").Formula = "=" & Q(3).Offset(0, i).AddressLocal
      Else
        shtPC.Range("CI_i").ClearContents
        shtPC.Range("CI_lambda_i").ClearContents
      End If
  End If
'==== End statistical tests about the eigenvalues ======='
  
  Set Q(4) = Q(3).Offset(2, 0)
  Q(4).Value = "Eigvec"
  'Q(4).Offset(0, 1).Resize(xcount, xcount).FormulaArray = "=MatEigenvector(" & Q(2).Offset(1, 1).Resize(xcount, xcount).AddressLocal & _
                                                          ",TRANSPOSE(" & Q(3).Offset(0, 1).Resize(1, xcount).AddressLocal & "))"

  Set Q(5) = Q(4).Offset(xcount + 2, 0)
  Q(5).Value = "mu"
  Q(5).Offset(0, 1).Resize(1, xcount).Value2 = mu
  
  Q(5).Offset(1, 0).Value = "sd"
  Q(5).Offset(1, 1).Resize(1, xcount).Value2 = sd
  
  Set Q(6) = Q(5).Offset(3, 0)
  Q(6).Value = "PC"
  x = Application.MMult(Data, Q(4).Offset(0, 1).Resize(xcount, xcount).Value2)
  For j = 1 To xcount
    For i = 1 To n
      x(i, j) = x(i, j) + mu(j)
    Next i
  Next j
  Q(6).Offset(1, 1).Resize(n, xcount).Value2 = x
  For j = 1 To xcount
    Q(6).Offset(0, j).Value = "PC" & j
  Next j
  
  shtPC.ListObjects.Add(xlSrcRange, Q(6).Offset(0, 1).Resize(n + 1, xcount), , xlYes).name = "PC"
  

'======== PC based on Correlation Matrix ===
'Not possible to make it a formula as columns may not be adjacent
'===== Equivalently, this is PC based on studentized data ===
  Set Q(12) = Q(2).Offset(0, xcount + 3)
  Q(12).Offset(-1, 0).Value = "PCA on Studentized data"
  Q(12).Value = "Cov(Correl)"
  For j = 1 To xcount
    Q(12).Offset(0, j).Value = headers(j)
    Q(12).Offset(j, 0).Value = headers(j)
  Next j
  
  Q(12).Offset(1, 1).Resize(xcount, xcount).Value2 = MCOVARIANCE(data_stu)
  
  Set Q(13) = Q(12).Offset(xcount + 3)
  Q(13).Value = "Eigval"
  Q(13).Offset(0, 1).Resize(2 + xcount, xcount).FormulaArray = "=CovEigenDecompQR(" & Q(12).Offset(1, 1).Resize(xcount, xcount).AddressLocal & ")"
  
'==== statistical tests about the eigenvalues ======='
  If shtPC.Range("tests_based_on").Value = "Correl" Then
      m = shtPC.Range("m").Value
      If m < 0 Or m >= xcount Then
        m = 0
        shtPC.Range("m").Value = 0
      End If
      Set alleigval = Q(13).Offset(0, 1).Resize(1, xcount)
      Set taileigval = Q(13).Offset(0, 1 + m).Resize(1, xcount - m)
      If m > 0 Then Set headeigval = Q(13).Offset(0, 1).Resize(1, m)
      shtPC.Range("total_var").Formula = "=SUM(" & alleigval.AddressLocal & ")"
      shtPC.Range("g_lambda").Formula = "=SUM(" & taileigval.AddressLocal & ") / total_var"
    
      If m = 0 Then
        shtPC.Range("var_pct").Value = ""
        shtPC.Range("aux").Formula = "= (1 - g_lambda) ^ 2 * SUMSQ(" & taileigval.AddressLocal & ")"
      Else
        shtPC.Range("var_pct").Formula = "=" & Q(13).Offset(0, m).AddressLocal & " / total_var"
        shtPC.Range("aux").Formula = "= g_lambda ^ 2 * SUMSQ(" & headeigval.AddressLocal & _
                              ") + (1 - g_lambda) ^ 2 * SUMSQ(" & taileigval.AddressLocal & ")"
      End If
      shtPC.Range("equal_small_eigval_chisq").FormulaArray = "=(n-1) *  ((p-m)*LN(SUM(" & taileigval.AddressLocal & _
                                                                   ")/(p-m)) - SUM(LN(" & taileigval.AddressLocal & ")))"
      
      i = shtPC.Range("CI_i").Value
      If i >= 1 And i <= xcount Then
        shtPC.Range("CI_lambda_i").Formula = "=" & Q(13).Offset(0, i).AddressLocal
      Else
        shtPC.Range("CI_i").ClearContents
        shtPC.Range("CI_lambda_i").ClearContents
      End If
  End If
'==== End statistical tests about the eigenvalues ======='
  
  Set Q(14) = Q(13).Offset(2, 0)
  Q(14).Value = "Eigvec"
  'Q(14).Offset(0, 1).Resize(xcount, xcount).FormulaArray = "=MatEigenvector(" & Q(12).Offset(1, 1).Resize(xcount, xcount).AddressLocal & _
                                                          ",TRANSPOSE(" & Q(13).Offset(0, 1).Resize(1, xcount).AddressLocal & "))"

  Set Q(15) = Q(14).Offset(xcount + 2, 0)
  Q(15).Value = "mu"
  Q(15).Offset(0, 1).Resize(1, xcount).Value = 0
  Q(15).Offset(1, 0).Value = "sd"
  Q(15).Offset(1, 1).Resize(1, xcount).Value = 1
  
  
  
  Set Q(16) = Q(15).Offset(3, 0)
  Q(16).Value = "PC Stu"
  x = Application.MMult(data_stu, Q(14).Offset(0, 1).Resize(xcount, xcount).Value2)
  For j = 1 To xcount
    For i = 1 To n
      x(i, j) = x(i, j)
    Next i
  Next j
  Q(16).Offset(1, 1).Resize(n, xcount).Value2 = x
  For j = 1 To xcount
    Q(16).Offset(0, j).Value = "PC_Stu" & j
  Next j
  
  shtPC.ListObjects.Add(xlSrcRange, Q(16).Offset(0, 1).Resize(n + 1, xcount), , xlYes).name = "PC_Stu"
  
  
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
  
  If Target.Row = shtPC.Range("m").Row And Target.Column = shtPC.Range("m").Column Then
    If existObject("DataArea_PC", "ListObject", shtPC) = True Then
      Call compute_PC
    End If
    GoTo lbl_exit
  End If
  
  If Target.Row = shtPC.Range("CI_i").Row And Target.Column = shtPC.Range("CI_i").Column Then
    If existObject("DataArea_PC", "ListObject", shtPC) = True Then
      Call compute_PC
    End If
    GoTo lbl_exit
  End If
  
  If Target.Row = shtPC.Range("tests_based_on").Row And Target.Column = shtPC.Range("tests_based_on").Column Then
    If existObject("DataArea_PC", "ListObject", shtPC) = True Then
      Call compute_PC
    End If
    GoTo lbl_exit
  End If
  
  If Target.Row = shtPC.Range("Left2Data").Row - 1 And _
     Target.Column >= shtPC.Range("DataArea_PC[#Headers]").Cells(1, 1).Column And _
     Target.Column <= shtPC.Range("DataArea_PC[#Headers]").Cells(1, 1).End(xlToRight).Column _
  Then
    
    Call compute_PC
    
    GoTo lbl_exit
  End If
lbl_exit:
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub
