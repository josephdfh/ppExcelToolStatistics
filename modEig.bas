Attribute VB_Name = "modEig"
Option Explicit

Function MatEigenvalue_max(A, Optional maxIter As Integer = 20)
On Error GoTo e
  If TypeName(A) = "Range" Then A = A.Value2
  Dim nrow&, ncol&, L1norm#, i&, itercount As Integer
  nrow = UBound(A) - LBound(A) + 1
  ncol = UBound(A, 1) - LBound(A, 1) + 1
  If nrow <> ncol Then GoTo e
  Dim diff, eigvec_old, eigvec
  ReDim eigvec(1 To nrow) As Double
  For i = LBound(eigvec) To UBound(eigvec)
    eigvec(i) = 1# / nrow
  Next i
  ReDim diff(1 To nrow) As Double
  Dim chg#: chg = 1
  itercount = 0
  Const tol As Double = 0.000000000000001
  While chg > tol And itercount < maxIter
    eigvec_old = eigvec
    eigvec = Application.MMult(eigvec, Application.Transpose(A))
    L1norm = Sqr(Application.SumSq(eigvec))
    For i = LBound(eigvec) To UBound(eigvec)
      eigvec(i) = eigvec(i) / L1norm
      diff(i) = eigvec(i) - eigvec_old(i)
    Next i
    chg = Sqr(Application.SumSq(diff))
    itercount = itercount + 1
  Wend
  If chg > tol Then
    MatEigenvalue_max = "#Not converged given 20 iterations in MatEigenvalue_max"
    Exit Function
  End If
  Dim eigval#, ans
  ans = Application.MMult(eigvec, Application.Transpose(A))
  For i = 0 To (UBound(eigvec) - LBound(eigvec))
    ans(LBound(ans) + i) = ans(LBound(ans) + i) / eigvec(LBound(eigvec) + i)
  Next i
  MatEigenvalue_max = Application.Median(ans)
  Exit Function
e: MatEigenvalue_max = "#Error in MatEigenvalue_max"
End Function
