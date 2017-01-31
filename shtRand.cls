VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtRand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Worksheet_Activate()
  On Error Resume Next
  With shtRand.Range("SelectDist").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Settings!" & shtSettings.Range("tblSettingsRand[Distribution]").AddressLocal
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  On Error Resume Next
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  If sameRange(Target, shtRand.Range("Generate")) Then
    Cancel = True
      
    Call generate
      
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtRand.Range("Clear")) Then
    Cancel = True
      
    Call shtRandClearStart
      
    GoTo lbl_exit
  End If
lbl_exit:
  Application.EnableEvents = True
  Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  On Error Resume Next
  Dim x
  If sameRange(Target, shtRand.Range("SelectDist")) Then
    Call prepareRand
    
  End If
End Sub

Private Sub prepareRand()
  On Error Resume Next
  Dim s
  For Each s In Array("DistPar1", "DistPar2", "DistPar3", "DistPar4", "DistPar5")
    shtRand.Range(CStr(s)).Value = ""
    shtRand.Range(CStr(s)).ClearComments
  Next s
  Dim distStr As String
  distStr = shtRand.Range("SelectDist").Value
  Dim npar As Integer: npar = vlookup2(shtSettings.ListObjects("tblSettingsRand"), distStr, "Distribution", "nPar").Value
  Dim i As Integer
  For i = 1 To npar
    shtRand.Range("DistPar" & i).Value = vlookup2(shtSettings.ListObjects("tblSettingsRand"), distStr, "Distribution", "Par" & i).Value
    shtRand.Range("DistPar" & i).AddComment Text:=vlookup2(shtSettings.ListObjects("tblSettingsRand"), distStr, "Distribution", "Par" & i & "Comm").Value
    
  Next i
End Sub

Private Sub generate()
  If Trim(LCase(shtRand.Range("SelectDist").Value)) = "multivariate normal" Then
    Call generate_Normal
    Exit Sub
  End If
  
End Sub

Private Sub generate_Normal()
  Dim x
  For Each x In shtRand.ChartObjects
    x.Delete
  Next x
  Dim n As Long: n = shtRand.Range("SampleSize").Value
  Dim mu As Range
  Set mu = shtRand.Range("DistPar1")
  If mu.End(xlDown).Row > 100 Then
    Call MsgBox("There should be an error in the parameter specifications.")
    Exit Sub
  End If
  Set mu = shtRand.Range(mu.Offset(1, 0), mu.End(xlDown))
  
  
  
  Dim d As Integer: d = mu.Rows.count
  Dim cov As Range: Set cov = shtRand.Range("DistPar2").Offset(1, 0).Resize(d, d)
  
  Dim muval: muval = mu.Value2
  Dim covval: covval = cov.Value2
    
  shtRand.Range("ClearStart").Value = mu.Cells(1, 1).AddressLocal
  Call shtRandClearStart
  
  mu.Value2 = muval
  cov.Value2 = covval
  
  Dim eigval As Range: Set eigval = shtRand.Range("DistPar2").Offset(1, d + 4).Resize(d, 1)
  eigval.FormulaArray = "=CovEigenDecompQR(" & cov.AddressLocal & ")"
  
  Dim eigvec As Range: Set eigvec = shtRand.Range("DistPar2").Offset(1, d + 6).Resize(d, d)
  eigvec.FormulaArray = "=MatEigenvector(" & cov.AddressLocal & ", " & eigval.AddressLocal & ")"
  
  Dim pc As Range: Set pc = eigvec.Cells(1, 1).Offset(d + 2, 0).Resize(n, d)
  pc.FormulaArray = "=NORMSINV(RAND()) * TRANSPOSE(SQRT(" & eigval.AddressLocal & "))"
  
  Dim res As Range: Set res = cov.Cells(1, 1).Offset(d + 2, 0).Resize(n, d)
  res.FormulaArray = "=MMULT(" & pc.AddressLocal & ", TRANSPOSE(" & eigvec.AddressLocal & ")) + TRANSPOSE(" & mu.AddressLocal & ")"
  
  res.Select
End Sub


Private Sub shtRandClearStart()
  Dim start As Range: Set start = shtRand.Range([clearstart].Value)
  Dim last As Range: Set last = shtRand.UsedRange: Set last = last.Cells(last.Rows.count, last.Columns.count)
  If last.Row >= start.Row And last.Column >= start.Column Then
    shtRand.Range(start, last).Clear
  End If
  Dim x
  For Each x In shtRand.ChartObjects
    x.Delete
  Next x
End Sub
