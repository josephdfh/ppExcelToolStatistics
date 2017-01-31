VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtCovTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  If sameRange(Target, shtCovTest.Range("paste_matrices")) Then
    Cancel = True
    
    Call pasteMatrices
    
    Exit Sub
  End If
  
  If sameRange(Target, shtCovTest.Range("clear")) Then
    Cancel = True
    
    Call covTestClear
    
    Exit Sub
  End If
  
  If sameRange(Target, shtCovTest.Range("covaddr")) Then
    Cancel = True
    
    Call getCovAddr
    
    Exit Sub
  End If
  
  If sameRange(Target, shtCovTest.Range("correladdr")) Then
    Cancel = True
    
    Call getCorrelAddr
    
    Exit Sub
  End If
  
  If sameRange(Target, shtCovTest.Range("LRT_First_Principle")) Or sameRange(Target, shtCovTest.Range("LRT_Improved")) Then
    Cancel = True
    
    Call LRT_Compute_CapitalLambda
    
    Exit Sub
  End If
  
  
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  On Error Resume Next
  Dim states(1 To 2) As Boolean
  states(1) = Application.EnableEvents
  states(2) = Application.ScreenUpdating
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  
  If sameRange(Target, shtCovTest.Range("selector")) Then
        Call viewSelector
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtCovTest.Range("g")) Then
    If shtCovTest.Range("g").Value > 0 Then
      shtCovTest.Range("p1_").Resize(1, shtCovTest.Range("g").Value).name = "'" & shtCovTest.name & "'!pis"
    Else
      shtCovTest.Range("p1_").Resize(1, 2).name = "'" & shtCovTest.name & "'!pis"
    End If
    
    GoTo lbl_exit
  End If


lbl_exit:
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub



Private Sub covTestClear()
  Dim p As Range: Set p = shtCovTest.Range("matrices")
  Dim last As Range: Set last = shtCovTest.UsedRange(shtCovTest.UsedRange.count)
  If last.Row >= p.Row Then
    shtCovTest.Range(p, last).ClearContents
  Else
    shtCovTest.Range(p, p.Offset(100, 1000)).ClearContents
  End If
  Dim x
  For Each x In Array("p", "n", "covaddr", "correladdr", "covhypoaddr", "g")
    shtCovTest.Range(CStr(x)).Value = ""
  Next x
  shtCovTest.Range(shtCovTest.Range("p1_"), shtCovTest.Range("p1_").Offset(0, 100)).Value = ""
  shtCovTest.Range(shtCovTest.Range("ln_corr1"), shtCovTest.Range("ln_corr1").Offset(0, 100)).Value = ""
End Sub

Private Sub getCorrelAddr()
  Dim R As Range: Set R = find2("Correlation Matrix", shtCovTest.Range("matrices"))
  Dim p As Integer: p = shtCovTest.Range("p").Value
  R.Offset(0, 1).Resize(p, p).name = "'" & shtCovTest.name & "'!correladdr"
End Sub

Private Sub getCovAddr()
  Dim R As Range: Set R = shtCovTest.Range("matrices")
  Dim p As Integer: p = shtCovTest.Range("p").Value
  R.Offset(0, 1).Resize(p, p).name = "'" & shtCovTest.name & "'!covAddr"
End Sub

Private Sub getCovHypoAddr()
  Dim R As Range: Set R = shtCovTest.Range("matrices").Offset(0, shtCovTest.Range("p").Value + 3)
  Dim p As Integer: p = shtCovTest.Range("p").Value
  R.Offset(0, 1).Resize(p, p).name = "'" & shtCovTest.name & "'!covhypoaddr"
End Sub

Private Sub LRT_Compute_CapitalLambda()
  Dim Rii As Range: Set Rii = shtCovTest.Range("correladdr").Cells(1, 1)
  Dim Q As Range: Set Q = shtCovTest.Range("ln_corr1")
  Dim PI As Range: Set PI = shtCovTest.Range("p1_")
  Dim g As Integer: g = shtCovTest.Range("g").Value
  Dim i As Integer
  For i = 1 To g
    Q.FormulaLocal = "=LN(MDETERM(" & Rii.Resize(PI.Value, PI.Value).AddressLocal & "))"
    Set Rii = Rii.Offset(PI, PI)
    Set PI = PI.Offset(0, 1) 'must be after the Rii line
    Set Q = Q.Offset(0, 1)
  Next i
  
End Sub

Private Sub pasteMatrices()
    Call covTestClear
    On Error GoTo lbl_exit
    Dim InputBoxRes: InputBoxRes = InputBox(prompt:="Sample Size (nobs) n=", Default:=shtPivot.Range("n").Value)
    If Trim(InputBoxRes) = "" Then Exit Sub
    InputBoxRes = CLng(InputBoxRes)
    shtCovTest.Range("n").Value = InputBoxRes
    Call shtCov.Range(shtCov.Range("leftto_cov_start"), shtCov.Range("correlAddr")).Copy
    shtCovTest.Range("matrices").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    shtCovTest.Range("p").Value = shtCov.Range("nrow").Value
    Call getCovAddr
    Call getCorrelAddr
    Call getCovHypoAddr
    shtCovTest.Range("matrices").Offset(0, shtCovTest.Range("p").Value + 3).Value = "Hypothesized Covariance"
    Application.GoTo reference:="matrices", Scroll:=True
    shtCovTest.Range("matrices").Offset(0, shtCovTest.Range("p").Value + 4).Select
lbl_exit:
End Sub

Private Sub viewSelector()
  On Error GoTo lbl_exit:
  Application.GoTo _
          reference:=find2(shtCovTest.Range("selector").Value, shtCovTest.Range("A1")), _
          Scroll:=True
lbl_exit:
End Sub
