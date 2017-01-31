VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtCancorr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
On Error GoTo lbl_exit
  If sameRange(Target, shtCancorr.Range("clear")) Then
    Cancel = True
    
    Call clear_cancorr
    Call clear_cancorr
    
    Exit Sub
  End If
  
  If sameRange(Target, shtCancorr.Range("Paste_data_from_Pivot")) Then
    Cancel = True
    
    Call Paste_data_from_Pivot
    
    Exit Sub
  End If
lbl_exit:
  Call enable
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  On Error GoTo lbl_exit
  
  If Target.Row = shtCancorr.Range("k").Row And Target.Column = shtCancorr.Range("k").Column Then
    If existObject("DataTable_CanCorr", "ListObject", shtCancorr) = True Then
      Call compute_simple_correl_matrix_for_all_v_and_w
    End If
    Exit Sub
  End If
  
  If Target.Row = shtCancorr.Range("Left_to_data_headers").Row - 1 And _
     Target.Column >= shtCancorr.Range("DataTable_CanCorr[#Headers]").Cells(1, 1).Column And _
     Target.Column <= shtCancorr.Range("DataTable_CanCorr[#Headers]").Cells(1, 1).End(xlToRight).Column _
  Then
      Call disable
      Call compute_simple_correl_matrix_for_all_v_and_w
      Call enable
    GoTo lbl_exit
  End If
lbl_exit:
  Call enable
End Sub



Sub clear_cancorr()
  On Error GoTo lbl_exit
  Dim states(1 To 2) As Boolean
  states(1) = Application.EnableEvents:  Application.EnableEvents = False
  states(2) = Application.ScreenUpdating: Application.ScreenUpdating = False
  Dim x
  For Each x In shtCancorr.ListObjects
    x.Delete
  Next x
  For Each x In shtCancorr.ChartObjects
    x.Delete
  Next x
  For Each x In Array("n", "vcount", "wcount")
    shtCancorr.Range(CStr(x)).ClearContents
  Next x
  shtCancorr.Range("k").Value = 0
  Dim R As Range: Set R = shtCancorr.Range("Left_to_data_headers").Offset(-1, 1)
  Dim last As Range: Set last = shtCancorr.UsedRange(shtCancorr.UsedRange.count)
  If last.Column >= R And last.Row >= R Then
    shtCancorr.Range(R, last).Clear
    shtCancorr.Range(R, last).ColumnWidth = 8.43
  Else
    R.Resize(100, 100).Clear
    R.Resize(100, 100).ColumnWidth = 8.43
  End If
  Set R = shtCancorr.Range("ZZ1").End(xlToRight).End(xlToLeft).Offset(0, 1)
  Set last = shtCancorr.UsedRange(shtCancorr.UsedRange.count)
  If last.Column >= R.Column Then
    shtCancorr.Range(R, last).EntireColumn.Delete
  End If
lbl_exit:
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub


Sub compute_simple_correl_matrix_for_all_v_and_w()
  On Error GoTo lbl_exit
  Dim states(1 To 2) As Boolean
  states(1) = Application.EnableEvents:  Application.EnableEvents = False
  states(2) = Application.ScreenUpdating: Application.ScreenUpdating = False
  
  
  
  Dim headers: headers = shtCancorr.Range("DataTable_CanCorr[#Headers]").Value2
  Dim n As Long: n = shtCancorr.Range("DataTable_CanCorr[#Data]").Columns(1).Cells.count
  shtCancorr.Range("n").Value = n
  Dim p As Long: p = shtCancorr.Range("DataTable_CanCorr[#Headers]").Cells.count
  Dim temp
  Dim vpos() As Long: ReDim vpos(1 To p) As Long
  Dim wpos() As Long: ReDim wpos(1 To p) As Long
  Dim vcount As Long:  vcount = 0
  Dim wcount As Long:  wcount = 0
  Dim i As Long, j As Long, k As Long
  Dim R As Range: Set R = shtCancorr.Range("Left_to_data_headers").Offset(-1, 0)
  R.Offset(0, p + 2).Value = "."
  shtCancorr.Range(R.Offset(0, p + 1), shtCancorr.UsedRange(shtCancorr.UsedRange.count)).Clear
  For j = 1 To p
    Set R = R.Offset(0, 1)
    If LCase(Trim(R.Value)) = "v" Then
      vcount = vcount + 1
      vpos(vcount) = j
    ElseIf LCase(Trim(R.Value)) = "w" Then
      wcount = wcount + 1
      wpos(wcount) = j
    End If
  Next j
  
  shtCancorr.Range("vcount").Value = vcount
  shtCancorr.Range("wcount").Value = wcount
  
  Dim Data: ReDim Data(1 To n, 1 To vcount + wcount) As Double
  Dim mu: ReDim mu(1 To vcount + wcount) As Double
  Dim sd: ReDim sd(1 To vcount + wcount) As Double
'=== prepare data : original data=='
  For j = 1 To vcount
    For i = 1 To n
      Data(i, j) = shtCancorr.Range("DataTable_CanCorr[#Data]").Cells(i, vpos(j)).Value
    Next i
    mu(j) = Application.Average(shtCancorr.Range("DataTable_CanCorr[#Data]").Columns(vpos(j)).Value2)
    sd(j) = Application.StDev(shtCancorr.Range("DataTable_CanCorr[#Data]").Columns(vpos(j)).Value2)
  Next j
  
  For j = vcount + 1 To vcount + wcount
    For i = 1 To n
      Data(i, j) = shtCancorr.Range("DataTable_CanCorr[#Data]").Cells(i, wpos(j - vcount)).Value
    Next i
    mu(j) = Application.Average(shtCancorr.Range("DataTable_CanCorr[#Data]").Columns(wpos(j - vcount)).Value2)
    sd(j) = Application.StDev(shtCancorr.Range("DataTable_CanCorr[#Data]").Columns(wpos(j - vcount)).Value2)
  Next j
  
'=== start output : on orginal data =='
  Dim Q(1 To 10) As Range
  Set Q(1) = shtCancorr.Range("Left_to_data_headers").Offset(0, 1).End(xlToRight).Offset(0, 3)
  Q(1).Value = "Covariance All"
  For i = 1 To vcount
    Q(1).Offset(0, i).Value = headers(1, vpos(i))
    Q(1).Offset(i, 0).Value = headers(1, vpos(i))
  Next i
  For i = 1 To wcount
    Q(1).Offset(0, vcount + i).Value = headers(1, wpos(i))
    Q(1).Offset(vcount + i, 0).Value = headers(1, wpos(i))
  Next i
  Q(1).Offset(1, 1).Resize(vcount + wcount, vcount + wcount).Value2 = MCOVARIANCE(Data)

'============ for V ===================
  Dim vdata: ReDim vdata(1 To n, 1 To vcount) As Double
  For i = 1 To n
    For j = 1 To vcount
      vdata(i, j) = Data(i, j)
    Next j
  Next i
  
  Set Q(2) = Q(1).Offset(vcount + wcount + 3, 0)
  Q(2).Value = "V QuadProd"
  Q(2).Offset(1, 0).Resize(vcount, vcount).FormulaArray = "=MMULT(MMULT(MINVERSE(" & Q(1).Offset(1, 1).Resize(vcount, vcount).AddressLocal & "), " & _
                                              Q(1).Offset(1, 1 + vcount).Resize(vcount, wcount).AddressLocal & "), MMULT(MINVERSE(" & _
                                              Q(1).Offset(1 + vcount, 1 + vcount).Resize(wcount, wcount).AddressLocal & "), " & Q(1).Offset(1 + vcount, 1).Resize(wcount, vcount).AddressLocal & "))"
  Set Q(3) = Q(2).Offset(0, vcount + 2)
  Q(3).Value = "Eigenvalues"
  
  
'  temp = MatEigenvalue_pow(q(2).Offset(1, 0).Resize(vcount, vcount).Value2)
'  q(3).Offset(1, 0).Resize(vcount, 1).Value2 = Application.Transpose(temp)
  If vcount = 1 Then
    Q(3).Offset(1, 0).Formula = "=" & Q(2).Offset(1, 0).AddressLocal
  Else
    Q(3).Offset(1, 0).Resize(vcount, 2 + vcount).FormulaArray = "=CovEigenDecompQR(" & Q(2).Offset(1, 0).Resize(vcount, vcount).AddressLocal & ",,,,True)"
  End If
  
'=== Wilk's Lambda: all eigenvalues will be the same so just use the first instances ==='
  k = shtCancorr.Range("k").Value
  If k < 0 Or k >= Application.Min(vcount, wcount) Then
    k = 0
    shtCancorr.Range("k").Value = k
  End If
  Set R = Q(3).Offset(1 + k, 0).Resize(Application.Min(vcount, wcount) - k, 1)
  shtCancorr.Range("Wilk_s_Lambda").FormulaArray = "=PRODUCT(1 - " & R.AddressLocal & ")"
  
  
  Q(3).Offset(0, -1).Value = "Cancorr"
  Q(3).Offset(1, -1).Resize(vcount, 1).FormulaArray = "=SQRT(" & Q(3).Offset(1, 0).Resize(vcount, 1).AddressLocal & ")"
  
  Set Q(4) = Q(3).Offset(0, 2)
  Q(4).Value = "V QuadProd Eigenvectors"
'  temp = MatEigenvector(q(2).Offset(1, 0).Resize(vcount, vcount).Value2, q(3).Offset(1, 0).Resize(vcount, 1).Value2)
'  q(4).Offset(1, 0).Resize(vcount, vcount).Value2 = temp
  If vcount = 1 Then
    Q(4).Offset(1, 0).Value = 1
  Else
    'Q(4).Offset(1, 0).Resize(vcount, vcount).FormulaArray = "=MatEigenvector(" & Q(2).Offset(1, 0).Resize(vcount, vcount).AddressLocal & ", " & Q(3).Offset(1, 0).Resize(vcount, 1).AddressLocal & ")"
  End If
  
  Set Q(5) = Q(4).Offset(0, vcount + 1)
  Q(5).Value = "sd along eigvecs"

  Q(5).Offset(1, 0).Resize(vcount, 1).FormulaArray = _
                      "=SQRT(GetDiagVector(MMULT(TRANSPOSE(" & _
                       Q(4).Offset(1, 0).Resize(vcount, vcount).AddressLocal & "),MMULT(" & _
                       Q(1).Offset(1, 1).Resize(vcount, vcount).AddressLocal & _
                       "," & Q(4).Offset(1, 0).Resize(vcount, vcount).AddressLocal & "))))"
  
  
  Set Q(6) = Q(5).Offset(0, 2)
  Q(6).Value = "studentized eigenvectors"
  For j = 1 To vcount
    Q(6).Offset(1, j - 1).Resize(vcount, 1).FormulaArray = "=" & Q(4).Offset(1, j - 1).Resize(vcount, 1).AddressLocal & " / " & Q(5).Offset(j, 0).AddressLocal
  Next j
  
'======compute canonically transformed original data specified in the v group====
  Set Q(7) = Q(1).Offset(-1, 2 * p + 8)
  Q(7).Value = "Canonical Transform on Original Data"
  For j = 1 To vcount
    Q(7).Offset(1, j - 1).Value = "v" & j
  Next j
  For j = 1 To wcount
    Q(7).Offset(1, vcount + j - 1).Value = "w" & j
  Next j
  Q(7).Offset(2, 0).Resize(n, vcount).Value2 = Application.MMult(vdata, Q(6).Offset(1, 0).Resize(vcount, vcount).Value2)
'============ for W ===================
  Dim wdata: ReDim wdata(1 To n, 1 To wcount) As Double
  For i = 1 To n
    For j = 1 To wcount
      wdata(i, j) = Data(i, vcount + j)
    Next j
  Next i
  
  Set Q(2) = Q(2).Offset(vcount + 3, 0)
  Q(2).Value = "W QuadProd"
  Q(2).Offset(1, 0).Resize(wcount, wcount).FormulaArray = "=MMULT(MMULT(MINVERSE(" & Q(1).Offset(1 + vcount, 1 + vcount).Resize(wcount, wcount).AddressLocal & "), " & _
                                                            Q(1).Offset(1 + vcount, 1).Resize(wcount, vcount).AddressLocal & "), MMULT(MINVERSE(" & _
                                                            Q(1).Offset(1, 1).Resize(vcount, vcount).AddressLocal & "), " & _
                                                            Q(1).Offset(1, 1 + vcount).Resize(vcount, wcount).AddressLocal & "))"
  Set Q(3) = Q(2).Offset(0, wcount + 2)
  Q(3).Value = "Eigenvalues"
  
  If wcount = 1 Then
    Q(3).Offset(1, 0).Formula = "=" & Q(2).Offset(1, 0).AddressLocal
  Else
    Q(3).Offset(1, 0).Resize(wcount, 2 + wcount).FormulaArray = "=CovEigenDecompQR(" & Q(2).Offset(1, 0).Resize(wcount, wcount).AddressLocal & ",,,,True)"
  End If
  
  Q(3).Offset(0, -1).Value = "Cancorr"
  Q(3).Offset(1, -1).Resize(wcount, 1).FormulaArray = "=SQRT(" & Q(3).Offset(1, 0).Resize(wcount, 1).AddressLocal & ")"
  
  Set Q(4) = Q(3).Offset(0, 2)
  Q(4).Value = "W QuadProd Eigenvectors"
  If wcount = 1 Then
    Q(4).Offset(1, 0).Value = 1
  Else
    'Q(4).Offset(1, 0).Resize(wcount, wcount).FormulaArray = "=MatEigenvector(" & Q(2).Offset(1, 0).Resize(wcount, wcount).AddressLocal & ", " & Q(3).Offset(1, 0).Resize(wcount, 1).AddressLocal & ")"
  End If
  
  Set Q(5) = Q(4).Offset(0, wcount + 1)
  Q(5).Value = "sd along eigvecs"
  
  Q(5).Offset(1, 0).Resize(wcount, 1).FormulaArray = _
                        "=SQRT(GetDiagVector(MMULT(TRANSPOSE(" & _
                         Q(4).Offset(1, 0).Resize(wcount, wcount).AddressLocal & "),MMULT(" & _
                         Q(1).Offset(vcount + 1, vcount + 1).Resize(wcount, wcount).AddressLocal & _
                         "," & Q(4).Offset(1, 0).Resize(wcount, wcount).AddressLocal & "))))"
  
  Set Q(6) = Q(5).Offset(0, 2)
  Q(6).Value = "studentized eigenvectors"
  For j = 1 To wcount
    Q(6).Offset(1, j - 1).Resize(wcount, 1).FormulaArray = "=" & Q(4).Offset(1, j - 1).Resize(wcount, 1).AddressLocal & " / " & Q(5).Offset(j, 0).AddressLocal
  Next j

'======compute canonically transformed original data specified in the w group====
  Q(7).Offset(2, vcount).Resize(n, wcount).Value2 = Application.MMult(wdata, Q(6).Offset(1, 0).Resize(wcount, wcount).Value2)
  shtCancorr.ListObjects.Add(xlSrcRange, Q(7).Offset(1, 0).Resize(n + 1, vcount + wcount), , xlYes).name = "CanonData"

'== prepare studentized data ===
  Dim data_studentized: ReDim data_studentized(1 To n, 1 To vcount + wcount) As Double
  For i = 1 To n
    For j = 1 To vcount + wcount
      data_studentized(i, j) = (Data(i, j) - mu(j)) / sd(j)
    Next j
  Next i


  Set Q(1) = Q(2).Offset(wcount + 5, 0)
  Q(1).Offset(-1, 0).Value = "On Studentized Data:"
  Q(1).Value = "Covaraince All"
  For i = 1 To vcount
    Q(1).Offset(0, i).Value = headers(1, vpos(i))
    Q(1).Offset(i, 0).Value = headers(1, vpos(i))
  Next i
  For i = 1 To wcount
    Q(1).Offset(0, vcount + i).Value = headers(1, wpos(i))
    Q(1).Offset(vcount + i, 0).Value = headers(1, wpos(i))
  Next i
  Q(1).Offset(1, 1).Resize(vcount + wcount, vcount + wcount).Value2 = MCOVARIANCE(data_studentized)
  
'============ for V ===================
  Dim vdata_studentized: ReDim vdata_studentized(1 To n, 1 To vcount) As Double
  For i = 1 To n
    For j = 1 To vcount
      vdata_studentized(i, j) = data_studentized(i, j)
    Next j
  Next i
  
  Set Q(2) = Q(1).Offset(vcount + wcount + 5, 0)
  Q(2).Value = "V QuadProd"
  Q(2).Offset(1, 0).Resize(vcount, vcount).FormulaArray = "=MMULT(MMULT(MINVERSE(" & Q(1).Offset(1, 1).Resize(vcount, vcount).AddressLocal & "), " & _
                                              Q(1).Offset(1, 1 + vcount).Resize(vcount, wcount).AddressLocal & "), MMULT(MINVERSE(" & _
                                              Q(1).Offset(1 + vcount, 1 + vcount).Resize(wcount, wcount).AddressLocal & "), " & Q(1).Offset(1 + vcount, 1).Resize(wcount, vcount).AddressLocal & "))"
  Set Q(3) = Q(2).Offset(0, vcount + 2)
  Q(3).Value = "Eigenvalues"
  
  If vcount = 1 Then
    Q(3).Offset(1, 0).Value = "=" & Q(2).Offset(1, 0).AddressLocal
  Else
    Q(3).Offset(1, 0).Resize(vcount, vcount + 2).FormulaArray = "=CovEigenDecompQR(" & Q(2).Offset(1, 0).Resize(vcount, vcount).AddressLocal & ",,,,True)"
  End If
  
  Q(3).Offset(0, -1).Value = "Cancorr"
  Q(3).Offset(1, -1).Resize(vcount, 1).FormulaArray = "=SQRT(" & Q(3).Offset(1, 0).Resize(vcount, 1).AddressLocal & ")"
  
  Set Q(4) = Q(3).Offset(0, 2)
  Q(4).Value = "V QuadProd Eigenvectors"
  
  If vcount = 1 Then
    Q(4).Offset(1, 0).Value = 1
  Else
    'Q(4).Offset(1, 0).Resize(vcount, vcount).FormulaArray = "=MatEigenvector(" & Q(2).Offset(1, 0).Resize(vcount, vcount).AddressLocal & ", " & Q(3).Offset(1, 0).Resize(vcount, 1).AddressLocal & ")"
  End If
  
  Set Q(5) = Q(4).Offset(0, vcount + 1)
  Q(5).Value = "sd along eigvecs"

  
  Q(5).Offset(1, 0).Resize(vcount, 1).FormulaArray = _
                      "=SQRT(GetDiagVector(MMULT(TRANSPOSE(" & _
                       Q(4).Offset(1, 0).Resize(vcount, vcount).AddressLocal & "),MMULT(" & _
                       Q(1).Offset(1, 1).Resize(vcount, vcount).AddressLocal & _
                       "," & Q(4).Offset(1, 0).Resize(vcount, vcount).AddressLocal & "))))"
  
  Set Q(6) = Q(5).Offset(0, 2)
  Q(6).Value = "studentized eigenvectors"
  For j = 1 To vcount
    Q(6).Offset(1, j - 1).Resize(vcount, 1).FormulaArray = "=" & Q(4).Offset(1, j - 1).Resize(vcount, 1).AddressLocal & " / " & Q(5).Offset(j, 0).AddressLocal
  Next j
  
'======compute canonically transformed original data specified in the v group====
  Set Q(8) = Q(7).Offset(0, vcount + wcount + 2)
  Q(8).Value = "Canonical Transform on Studentized Data"
  For j = 1 To vcount
    Q(8).Offset(1, j - 1).Value = "v" & j
  Next j
  For j = 1 To wcount
    Q(8).Offset(1, vcount + j - 1).Value = "w" & j
  Next j
  
  Q(8).Offset(2, 0).Resize(n, vcount).Value2 = Application.MMult(vdata_studentized, Q(6).Offset(1, 0).Resize(vcount, vcount).Value2)
'============ for W ===================
  Dim wdata_studentized: ReDim wdata_studentized(1 To n, 1 To wcount) As Double
  For i = 1 To n
    For j = 1 To wcount
      wdata_studentized(i, j) = data_studentized(i, vcount + j)
    Next j
  Next i
  
  Set Q(2) = Q(2).Offset(vcount + 3, 0)
  Q(2).Value = "W QuadProd"
  Q(2).Offset(1, 0).Resize(wcount, wcount).FormulaArray = "=MMULT(MMULT(MINVERSE(" & Q(1).Offset(1 + vcount, 1 + vcount).Resize(wcount, wcount).AddressLocal & "), " & _
                                                            Q(1).Offset(1 + vcount, 1).Resize(wcount, vcount).AddressLocal & "), MMULT(MINVERSE(" & _
                                                            Q(1).Offset(1, 1).Resize(vcount, vcount).AddressLocal & "), " & _
                                                            Q(1).Offset(1, 1 + vcount).Resize(vcount, wcount).AddressLocal & "))"
  Set Q(3) = Q(2).Offset(0, wcount + 2)
  Q(3).Value = "Eigenvalues"
  
  If wcount = 1 Then
    Q(3).Offset(1, 0).Formula = "=" & Q(2).Offset(1, 0).AddressLocal
  Else
    Q(3).Offset(1, 0).Resize(wcount, wcount + 2).FormulaArray = "=CovEigenDecompQR(" & Q(2).Offset(1, 0).Resize(wcount, wcount).AddressLocal & ",,,,True)"
  End If
  
  Q(3).Offset(0, -1).Value = "Cancorr"
  Q(3).Offset(1, -1).Resize(wcount, 1).FormulaArray = "=SQRT(" & Q(3).Offset(1, 0).Resize(wcount, 1).AddressLocal & ")"
  
  Set Q(4) = Q(3).Offset(0, 2)
  Q(4).Value = "W QuadProd Eigenvectors"
  
  If wcount = 1 Then
    Q(4).Offset(1, 0).Value = 1
  Else
    'Q(4).Offset(1, 0).Resize(wcount, wcount).FormulaArray = "=MatEigenvector(" & Q(2).Offset(1, 0).Resize(wcount, wcount).AddressLocal & ", " & Q(3).Offset(1, 0).Resize(wcount, 1).AddressLocal & ")"
  End If
  
  Set Q(5) = Q(4).Offset(0, wcount + 1)
  Q(5).Value = "sd along eigvecs"

  Q(5).Offset(1, 0).Resize(wcount, 1).FormulaArray = _
                      "=SQRT(GetDiagVector(MMULT(TRANSPOSE(" & _
                       Q(4).Offset(1, 0).Resize(wcount, wcount).AddressLocal & "),MMULT(" & _
                       Q(1).Offset(vcount + 1, vcount + 1).Resize(wcount, wcount).AddressLocal & _
                       "," & Q(4).Offset(1, 0).Resize(wcount, wcount).AddressLocal & "))))"
  
  Set Q(6) = Q(5).Offset(0, 2)
  Q(6).Value = "studentized eigenvectors"
  For j = 1 To wcount
    Q(6).Offset(1, j - 1).Resize(wcount, 1).FormulaArray = "=" & Q(4).Offset(1, j - 1).Resize(wcount, 1).AddressLocal & " / " & Q(5).Offset(j, 0).AddressLocal
  Next j
'======compute canonically transformed original data specified in the w group====
  Q(8).Offset(2, vcount).Resize(n, wcount).Value2 = Application.MMult(wdata_studentized, Q(6).Offset(1, 0).Resize(wcount, wcount).Value2)
  shtCancorr.ListObjects.Add(xlSrcRange, Q(8).Offset(1, 0).Resize(n + 1, vcount + wcount), , xlYes).name = "CanonDataStudentized"
  
  Set Q(10) = Q(8).Offset(0, vcount + wcount + 2)
  Q(10).Value = "Full Correl"
  For j = 1 To vcount
    Q(10).Offset(0, j).Value = headers(1, vpos(j))
    Q(10).Offset(j, 0).Value = headers(1, vpos(j))
    Q(10).Offset(0, vcount + wcount + j).Value = "v" & j
    Q(10).Offset(vcount + wcount + j, 0).Value = "v" & j
  Next j
  
  For j = 1 To wcount
    Q(10).Offset(0, vcount + j).Value = headers(1, wpos(j))
    Q(10).Offset(vcount + j, 0).Value = headers(1, wpos(j))
    Q(10).Offset(0, 2 * vcount + wcount + j).Value = "w" & j
    Q(10).Offset(2 * vcount + wcount + j, 0).Value = "w" & j
  Next j
  
  Dim bigdata: ReDim bigdata(1 To n, 1 To 2 * (vcount + wcount)) As Double
  For i = 1 To n
    For j = 1 To vcount
      bigdata(i, j) = Data(i, j)
      bigdata(i, vcount + wcount + j) = shtCancorr.Range("CanonData[#Data]").Cells(i, j).Value
    Next j
    For j = 1 To wcount
      bigdata(i, vcount + j) = Data(i, vcount + j)
      bigdata(i, 2 * vcount + wcount + j) = shtCancorr.Range("CanonData[#Data]").Cells(i, vcount + j).Value
    Next j
  Next i
  
  Q(10).Offset(1, 1).Resize(2 * (vcount + wcount), 2 * (vcount + wcount)).Value2 = MCORRELATION(bigdata)
  
lbl_exit:
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub

Sub Paste_data_from_Pivot()
  Call clear_cancorr
  On Error GoTo lbl_exit
  shtPivot.Range("DataArea[#All]").Copy
  shtCancorr.Range("Left_to_data_headers").Offset(0, 1).PasteSpecial Paste:=xlPasteAll
  shtCancorr.Range("Left_to_data_headers").Offset(0, 1).ListObject.name = "DataTable_CanCorr"
  Application.CutCopyMode = False
lbl_exit:
End Sub

