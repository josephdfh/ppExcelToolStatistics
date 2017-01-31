VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtMultiCov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  If sameRange(Target, shtMultiCov.Range("startdata")) Then
    Cancel = True
    
    Call pave
    
    Exit Sub
  End If
  
  If sameRange(Target, shtMultiCov.Range("fetch_pivot")) Then
    Cancel = True
    
    Call fetch_pivot
    
    Exit Sub
  End If
  
  
  If sameRange(Target, shtMultiCov.Range("clear")) Then
    Cancel = True
    
    Call multicovClear
    
    Exit Sub
  End If
  
End Sub

Private Sub pave()
  
  Dim p As Integer: p = shtMultiCov.Range("p").Value
  Dim n As Integer: n = shtMultiCov.Range("n").Value
  Dim k As Integer: k = shtMultiCov.Range("k").Value
  If k < 1 Or p < 1 Then
    Exit Sub
  End If
  
  Dim i As Integer
  Dim R As Range
  Set R = shtMultiCov.Range("startdata")
  shtMultiCov.Range(R.Offset(1, 0), R.Offset(50000, 0).End(xlUp).Offset(30, p * 5)).Value = ""
  
  
  
  Dim fmla(1 To 10) As String
  
  For i = 1 To k
    Set R = R.Offset(1, 0)
    R.Value = "n" & i
    R.Offset(0, 1).Value = shtMultiCov.Range("n1_").Offset(0, i - 1).Value
    R.Offset(1, 0).Value = "start S" & i
    R.Offset(p, 0).Value = "end S" & i
    R.Offset(0, p + 2).Formula = "=ABS(MDETERM(" & R.Offset(1, 1).Resize(p, p).AddressLocal & "))"
    R.Offset(0, p).Formula = "=" & R.Offset(0, 1).AddressLocal & " * LN(" & R.Offset(0, p + 2).AddressLocal & " * ((" & R.Offset(0, 1).AddressLocal & _
                             " - 1) / " & R.Offset(0, 1).AddressLocal & ")^p)"
    R.Offset(0, p + 1).Formula = "=(" & R.Offset(0, 1).AddressLocal & " - 1) * LN(" & R.Offset(0, p + 2).AddressLocal & ")"
    fmla(1) = fmla(1) & " + (" & shtMultiCov.Range("n1_").Offset(0, i - 1).AddressLocal & " - 1) * " & _
                      R.Offset(1, 1).Resize(p, p).AddressLocal
    fmla(2) = fmla(2) & " + " & R.Offset(0, p).AddressLocal
    fmla(3) = fmla(3) & " + " & R.Offset(0, p + 1).AddressLocal
    fmla(4) = fmla(4) & " + " & R.Offset(1, 1).Resize(p, p).AddressLocal & " / " & shtMultiCov.Range("n1_").Offset(0, i - 1).AddressLocal
    Set R = R.Offset(p + 2)
  Next i
  fmla(1) = "=(" & fmla(1) & ") / n"
  fmla(2) = "=" & fmla(2)
  fmla(3) = "=" & fmla(3)
  Set R = shtMultiCov.Range("startdata").Offset(1, 4 + p)
  R.Value = "Sigmahat (CSSP divided by n)"
  R.Offset(1, 0).Resize(p, p).FormulaArray = fmla(1)
  shtMultiCov.Range("sum_niLnSighat").Formula = fmla(2)
  shtMultiCov.Range("sum_ni_1_LnSi").Formula = fmla(3)
  shtMultiCov.Range("nLn_Sighat").Formula = "=n * LN(MDETERM(" & R.Offset(1, 0).Resize(p, p).AddressLocal & "))"
  
  Dim r1 As Range
  Set r1 = R.Offset(p + 2, 0)
  r1.Value = "Pooled Sample Covariances (CSSP divided by n - k)"
  With r1.Offset(1, 0).Resize(p, p)
    .FormulaArray = "= n / (n - k) *" & R.Offset(1, 0).Resize(p, p).AddressLocal
    .name = "'" & shtMultiCov.name & "'!S"
  End With
  shtMultiCov.Range("LnDetS").Formula = "=LN(MDETERM(" & r1.Offset(1, 0).Resize(p, p).AddressLocal & "))"
  
  shtMultiCov.Range("xbar1").Clear
  With shtMultiCov.Range("Mean1").Offset(0, 1).Resize(1, p)
    .name = "'" & shtMultiCov.name & "'!xbar1"
    .Style = "Input"
    .HorizontalAlignment = xlCenter
    .NumberFormat = "0.00"
    .Font.Bold = True
    .Font.Color = RGB(0, 0, 0)
  End With
  
  shtMultiCov.Range("xbar2").Clear
  With shtMultiCov.Range("Mean2").Offset(0, 1).Resize(1, p)
    .name = "'" & shtMultiCov.name & "'!xbar2"
    .Style = "Input"
    .HorizontalAlignment = xlCenter
    .NumberFormat = "0.00"
    .Font.Bold = True
    .Font.Color = RGB(0, 0, 0)
  End With
  
  shtMultiCov.Range("diff").Clear
  With shtMultiCov.Range("Left2Diff").Offset(0, 1).Resize(1, p)
    .name = "'" & shtMultiCov.name & "'!diff"
    .Style = "Input"
    .HorizontalAlignment = xlCenter
    .Font.Bold = True
    .Font.Color = RGB(0, 0, 0)
  End With
  
  Dim r2 As Range
  Set r2 = r1.Offset(p + 2, 0)
  r2.Value = "U Matrix (Sum of Sk / nk)"
  With r2.Offset(1, 0).Resize(p, p)
    .FormulaArray = "=" & fmla(4)
    .name = "'" & shtMultiCov.name & "'!U"
  End With
  '=MMULT(xbar1-xbar2-Diff,MMULT(MINVERSE(U),TRANSPOSE(xbar1-xbar2-Diff))) / (MMULT(xbar1-xbar2-Diff,MMULT(MMULT(MINVERSE(U), MMULT(B11:E14/E2, MINVERSE(U))),TRANSPOSE(xbar1-xbar2-Diff))) / (E2-1) + MMULT(xbar1-xbar2-Diff,MMULT(MMULT(MINVERSE(U), MMULT(B18:E21/F2, MINVERSE(U))),TRANSPOSE(xbar1-xbar2-Diff))) / (F2 - 1))
  shtMultiCov.Range("aux2_").FormulaArray = "=MMULT(xbar1-xbar2-Diff,MMULT(MMULT(MINVERSE(U), MMULT(" & _
        shtMultiCov.Range("startdata").Offset(2, 1).Resize(p, p).AddressLocal & _
        "/n1_, MINVERSE(U))),TRANSPOSE(xbar1-xbar2-Diff))) / (n1_-1)"
  
  shtMultiCov.Range("aux3_").FormulaArray = "=MMULT(xbar1-xbar2-Diff,MMULT(MMULT(MINVERSE(U), MMULT(" & _
        shtMultiCov.Range("startdata").Offset(5 + p, 1).Resize(p, p).AddressLocal & _
        "/n2_, MINVERSE(U))),TRANSPOSE(xbar1-xbar2-Diff))) / (n2_ - 1)"
  
End Sub

Private Sub fetch_pivot()
  
  Dim R As Range: Set R = shtPivot.Range("above_data_start").Offset(1 + shtPivot.Range("n").Value, 0).End(xlDown)
  Dim n As Integer: n = R.Offset(0, 1).Value
  
  Dim r2 As Range: Set r2 = R.Offset(10000, 0).End(xlUp)
  
  Dim p As Integer: p = shtPivot.Range(R.Offset(2, 1), R.Offset(2, 0).End(xlToRight)).Columns.count
  If r2.Row - R.Row < p + 2 Then Exit Sub  'only 1 matrix
  
  Dim r3 As Range: Set r3 = R.End(xlDown).End(xlDown)
  Dim step As Integer: step = r3.Row - R.Row
  Dim k As Integer: k = Int((r2.Row - R.Row) / step) + 1
  shtMultiCov.Range("k").Value = k
  
  shtMultiCov.Range("p").Value = p
  
  Call pave
  Dim Q As Range: Set Q = shtMultiCov.Range("startdata").Offset(1, 0)
  
  
  Dim i As Integer
  For i = 1 To k
    shtMultiCov.Range("n1_").Offset(0, i - 1).Value = R.Offset(0, 1).Value
    Q.Offset(0, 1).Value = R.Offset(0, 1).Value
    Q.Offset(1, 1).Resize(p, p).Value2 = R.Offset(2, 1).Resize(p, p).Value2
    Set R = R.Offset(p + 1, 0).End(xlDown)
    Set Q = Q.Offset(p, 0).End(xlDown)
  Next i
  shtMultiCov.Range("n").Value = shtMultiCov.Range("sum_ni").Value
End Sub

Private Sub multicovClear()
  Dim R As Range: Set R = shtMultiCov.Range("n1_")
  shtMultiCov.Range(R, R.Offset(0, Application.Max(10, shtMultiCov.Range("k").Value))).Value = ""
  Dim last As Range: Set last = shtMultiCov.UsedRange(shtMultiCov.UsedRange.count)
  Set R = shtMultiCov.Range("startdata").Offset(1, 0)
  If last.Row >= R.Row Then
    shtMultiCov.Range(last, R).Clear
  Else
    shtMultiCov.Range(R, R.Offset(Application.Max(100, shtMultiCov.Range("p").Value * 10), shtMultiCov.Range("p").Value * 7)).Clear
  End If
  Dim names:  names = Array("k", "p", "n", "nLn_Sighat", "sum_niLnSighat", "LnDetS", "sum_ni_1_LnSi")
  Dim x
  For Each x In names
    shtMultiCov.Range(CStr(x)).Value = ""
  Next x
  
  names = Array("xbar1", "xbar2", "diff")
  For Each x In names
    shtMultiCov.Range(CStr(x)).Clear
  Next x
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  On Error Resume Next
  Dim states(1 To 2) As Boolean
  states(1) = Application.EnableEvents
  states(2) = Application.ScreenUpdating
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  
  If sameRange(Target, shtMultiCov.Range("k")) Then
    
    If shtMultiCov.Range("k").Value > 0 Then
      shtMultiCov.Range("n1_").Resize(1, shtMultiCov.Range("k").Value).name = "'" & shtMultiCov.name & "'!nks"
    Else
      shtMultiCov.Range("n1_").Resize(1, 2).name = "'" & shtMultiCov.name & "'!nks"
    End If
    
    GoTo lbl_exit
  End If


lbl_exit:
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub
