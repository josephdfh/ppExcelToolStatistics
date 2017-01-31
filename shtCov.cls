VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtCov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  If sameRange(Target, shtCov.Range("getdim")) Then
    Cancel = True
    
    Call getDim(copyMatrix:=True)
    
    Exit Sub
  End If
  
  If sameRange(Target, shtCov.Range("getCorrel")) Then
    Cancel = True
    
    Call getCorrel
    
    Exit Sub
  End If
  
  If sameRange(Target, shtCov.Range("clear")) Then
    Cancel = True
    
    Call covarianceClear
    
    Exit Sub
  End If
  
  If sameRange(Target, shtCov.Range("complete_Covariance")) Then
    Cancel = True
    
    Call completeCovariance
    
    Exit Sub
  End If
  
  If sameRange(Target, shtCov.Range("copyall")) Then
    Cancel = True
    
    Call copyall
    
    Exit Sub
  End If
  
  If sameRange(Target, shtCov.Range("PlotCov")) Then
    Cancel = True
    
    Call PlotCov
    
    Exit Sub
  End If
  
  If sameRange(Target, shtCov.Range("Clear_Paste_Split")) Then
    Cancel = True
    
    Call clear_paste_split
    
    Exit Sub
  End If
  
End Sub

Private Sub clear_paste_split()
  Dim R As Range
  Dim res As Boolean: res = pasteCurrentClipboardToShtTemp(R)
  If res = False Then Exit Sub
  
  Call covarianceClear
  
  R.Copy
  shtCov.Range("leftto_cov_start").Offset(0, 1).PasteSpecial Paste:=xlValues
  Call completeCovariance
  Call getCorrel
  
  
End Sub

Private Sub completeCovariance()
  On Error GoTo lbl_exit
  Call getDim(copyMatrix:=False)
  Dim n As Integer
  n = CInt(InputBox(prompt:="Pls enter/confirm the dimension of the covariance matrix:", Default:=shtCov.Range("nrow").Value))
  If n < 2 Then Exit Sub
  Dim mat As Range: Set mat = shtCov.Range("leftto_cov_start").Offset(0, 1).Resize(n, n)
  Dim i As Integer, j As Integer
  For i = 2 To n Step 1
    For j = 1 To (i - 1) Step 1
      If TypeName(mat(i, j).Value) = "String" And TypeName(mat(j, i).Value) = "Double" Then
          mat(i, j).Value = mat(j, i).Value
      ElseIf TypeName(mat(j, i).Value) = "String" And TypeName(mat(i, j).Value) = "Double" Then
          mat(j, i).Value = mat(i, j).Value
      End If
    Next j
  Next i
  Call highlight_Nonumer_Cells(mat)
lbl_exit:
End Sub

Private Sub copyall()
  On Error GoTo lbl_exit:
  Call shtCov.Range(shtCov.Range("leftto_cov_start"), shtCov.Range(shtCov.Range("correlAddr").Value)).Copy
lbl_exit:
End Sub


Private Sub covarianceClear()
  Dim last As Range: Set last = shtCov.UsedRange(shtCov.UsedRange.count)
  Dim p As Range: Set p = shtCov.Range("leftto_cov_start")
  If last.Row >= p.Row And last.Column >= p.Column Then
    shtCov.Range(p, last).Clear
  End If
  Dim x
  For Each x In shtCov.ListObjects: x.Delete: Next x
  For Each x In shtCov.ChartObjects: x.Delete: Next x
  
  Dim params As Variant: params = Array("nrow", "ncol")
  For Each x In params: shtCov.Range(x).Value = "": Next x
  
  p.Value = "Covariance Matrix"
  p.Offset(0, 1).Value = "paste here"
  p.Offset(0, 1).Select
End Sub

Private Sub fillErrorInDimensionCells()
  shtCov.Range("nrow").Value = "Error"
  shtCov.Range("ncol").Value = "Error"
End Sub


Private Sub getCorrel()
  Call getDim
  If havingDimensionError = True Then
    MsgBox "Dimsionsion error exists --> Abort operation"
    Exit Sub
  End If
  
  Dim p As Range:  Set p = shtCov.Range("leftto_cov_start")
  Dim n As Integer: n = shtCov.Range("nrow").Value
  
  Dim Q As Range: Set Q = p.Offset(n + 1, 1)
  Q.Offset(-1, 0).Resize(n + 2, n + 5).Clear
  
  Q.Resize(n, 1).FormulaArray = "=GetDiagVector(" & p.Offset(0, 1).Resize(n, n).AddressLocal & ")"
  Q.Offset(0, -1).Value = "Var & sd"
  Q.Offset(0, 1).Resize(n, 1).FormulaArray = "=sqrt(" & Q.Resize(n, 1).AddressLocal & ")"

  Q.Offset(n + 1, 0).Resize(n, n).FormulaArray = "=MMULT(MMULT(MakeDiagMatrix(1 / " & Q.Offset(0, 1).Resize(n, 1).AddressLocal & "), " & _
                                                         p.Offset(0, 1).Resize(n, n).AddressLocal & "), " & _
                                                   "MakeDiagMatrix(1 / " & Q.Offset(0, 1).Resize(n, 1).AddressLocal & "))"
  Q.Offset(n + 1, -1).Value = "Correlation Matrix"
  Q.Offset(n + 1, 0).Resize(n, n).name = "'" & shtCov.name & "'!correlAddr"
  
  'eigen
  Dim R(1 To 4) As Range
  Set R(1) = Q.Offset(2 * n + 2, -1)
  R(1).Value = "Cov Eigenvalue"
  R(1).Offset(0, 1).Resize(n + 2, n).FormulaArray = "=CovEigenDecompQR(" & _
                                                p.Offset(0, 1).Resize(n, n).AddressLocal & ")"
  Set R(2) = R(1).Offset(2, 0)
  R(2).Value = "Cov Eigenvector"
  
  Set R(3) = R(2).Offset(n + 1, 0)
  R(3).Value = "Corr Eigenvalue"
  R(3).Offset(0, 1).Resize(n + 2, n).FormulaArray = "=CovEigenDecompQR(" & _
                                                Q.Offset(n + 1, 0).Resize(n, n).AddressLocal & ")"
  Set R(4) = R(3).Offset(2, 0)
  R(4).Value = "Corr Eigenvector"
End Sub

Private Sub getDim(Optional ByVal copyMatrix As Boolean = False)
  Dim p As Range
  Set p = shtCov.Range("leftto_cov_start")
  If Trim(p.Offset(0, 1).Value) = "" Or _
     Trim(p.Offset(1, 1).Value) = "" Or _
     Trim(p.Offset(0, 2).Value) = "" Or _
     Trim(p.Offset(1, 2).Value) = "" _
  Then
    Call fillErrorInDimensionCells
    Exit Sub
  End If
  
  Dim nrow As Integer, ncol As Integer
  nrow = p.Offset(0, 1).End(xlDown).Row - p.Row + 1
  ncol = p.Offset(0, 1).End(xlToRight).Column - p.Column
  Call setDimensionCells(nrow, ncol)
  If havingDimensionError Then Exit Sub
  
  Call selectCovarianceMatrix
  Selection.name = "'" & shtCov.name & "'!covAddr"
  If copyMatrix Then Selection.Copy
End Sub



Private Function havingDimensionError() As Boolean
  havingDimensionError = True
  If shtCov.Range("nrow").Value = "Error" Or shtCov.Range("ncol").Value = "Error" Then
    Exit Function
  End If
  
  If TypeName(shtCov.Range("nrow").Value) <> "Double" Or TypeName(shtCov.Range("ncol").Value) <> "Double" Then
    Exit Function
  End If
  
  If CLng(shtCov.Range("nrow").Value) <> shtCov.Range("ncol").Value Then
    Exit Function
  End If
  
  If CLng(shtCov.Range("ncol").Value) <> shtCov.Range("nrow").Value Then
    Exit Function
  End If
  
  If shtCov.Range("nrow").Value < 2 Then
    Exit Function
  End If
  
  'finally passed --> no dimension error now
  havingDimensionError = False
End Function


Private Sub selectCorrelationMatrix()
  Dim correlAddress As String: correlAddress = shtCov.Range("correlAddr").Value
  shtCov.Range(correlAddress).Select
End Sub

Private Sub selectCovarianceMatrix()
  Dim p As Range:  Set p = shtCov.Range("leftto_cov_start")
  Dim n As Integer: n = shtCov.Range("nrow").Value
  If Not havingDimensionError() Then
    p.Offset(0, 1).Resize(n, n).Select
  End If
End Sub


Private Sub setDimensionCells(ByVal nrow As Integer, ByVal ncol As Integer)
  If nrow >= 1 Then
    shtCov.Range("nrow").Value = nrow
  Else
    shtCov.Range("nrow").Value = "Error"
  End If
  
  If nrow >= 1 Then
    shtCov.Range("ncol").Value = ncol
  Else
    shtCov.Range("ncol").Value = "Error"
  End If
  
End Sub

Private Sub PlotCov()
  
  On Error Resume Next
  Dim addrStr As String
  Cov2x2Picker.Show
  addrStr = Cov2x2Picker.covAddr
  If Range(addrStr).Cells.count <> 4 Then
    MsgBox ("There is an error in the selected cells.")
    Exit Sub
  End If
  Dim R As Range
  Dim x(1 To 4) As Double
  
  Dim i As Integer
  i = 1
  For Each R In Range(addrStr)
    x(i) = R.Value
    i = i + 1
  Next R
  
  Dim cov(1 To 2, 1 To 2) As Double
  cov(1, 1) = x(1)
  cov(1, 2) = x(2)
  cov(2, 1) = x(3)
  cov(2, 2) = x(4)
  
  Call EllipsePlotCov(cov)
End Sub
