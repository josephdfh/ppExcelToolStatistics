Attribute VB_Name = "modEig"
Option Explicit
Function norm(v)
  norm = Sqr(Application.SumSq(v))
End Function
Public Function GetDiagVector(ByVal mat, Optional ByVal intoColumn As Boolean = True)
  If TypeName(mat) = "Range" Then
    mat = mat.Value2
  End If
  Dim n As Integer
  n = Application.Min(UBound(mat, 1) - LBound(mat, 1) + 1, UBound(mat, 2) - LBound(mat, 2) + 1)
  Dim i As Integer
  
  Dim res(): ReDim res(1 To n)
  For i = 1 To n
    res(i) = mat(LBound(mat, 1) + i - 1, LBound(mat, 2) + i - 1)
  Next i
  If intoColumn Then
    res = Application.Transpose(res)
  End If
  GetDiagVector = res
End Function
Function QR(A, Optional Q = Null, Optional R)
On Error GoTo e
  If TypeName(A) = "Range" Then A = A.Value2
  Dim nrow&, ncol&
  Dim i&, j&, k&, y, ans, res
  nrow = UBound(A, 1) - LBound(A, 1) + 1
  ncol = UBound(A, 2) - LBound(A, 2) + 1
  If nrow < ncol Then
    QR = "#Error: nrow must >= ncol for QR"
    Exit Function
  End If
  If IsNull(Q) Then
    ReDim Q(1 To nrow, 1 To ncol) As Double
  End If
  For k = 1 To nrow
    Q(k, 1) = A(LBound(A, 1) + k - 1, LBound(A, 2))
  Next k
  ReDim R(1 To ncol, 1 To ncol) As Double
  ReDim res(1 To (nrow + ncol + 1), 1 To ncol)
  For j = 1 To ncol
    res(nrow + 1, j) = ""
    ReDim y(1 To nrow) As Double
    For k = 1 To nrow: y(k) = A(k, j): Next k
    For i = 1 To (j - 1)
      ans = 0
      For k = 1 To nrow
        ans = ans + Q(k, i) * y(k)
      Next k
      R(i, j) = ans: res(nrow + i + 1, j) = R(i, j)
      For k = 1 To nrow
        y(k) = y(k) - R(i, j) * Q(k, i)
      Next k
    Next i
    R(j, j) = norm(y): res(nrow + j + 1, j) = R(i, j)
    For k = 1 To nrow
      Q(k, j) = y(k) / R(j, j): res(k, j) = Q(k, j)
    Next k
  Next j
  QR = res
e:
End Function
Function MatEigenvalue_max(A, Optional maxiter As Integer = 100)
On Error GoTo e
  If TypeName(A) = "Range" Then A = A.Value2
  Dim nrow&, ncol&, L2norm#, i&, itercount As Integer, eigval#
  nrow = UBound(A, 1) - LBound(A, 1) + 1
  ncol = UBound(A, 2) - LBound(A, 2) + 1
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
  While chg > tol And itercount < maxiter
    eigvec_old = eigvec
    eigvec = Application.MMult(eigvec, Application.Transpose(A))
    eigval = Application.SumProduct(eigvec, eigvec_old)  'Rayleigh
    L2norm = norm(eigvec)
    For i = LBound(eigvec) To UBound(eigvec)
      eigvec(i) = eigvec(i) / L2norm
      diff(i) = eigvec(i) - eigvec_old(i)
    Next i
    chg = norm(diff)
    itercount = itercount + 1
  Wend
  If chg > tol Then
    MatEigenvalue_max = "#Not converged given 20 iterations in MatEigenvalue_max"
    Exit Function
  End If
  MatEigenvalue_max = eigval
  Exit Function
e: MatEigenvalue_max = "#Error in MatEigenvalue_max"
End Function

Function CovEigenDecompQR(A, Optional maxiter = 1000, Optional eigvec, Optional eigval)
  On Error GoTo e
  If TypeName(A) = "Range" Then A = A.Value2
  Dim nrow&, ncol&, L2norm#, itercount As Integer, res, i&, j&, k&, chg, ans
  nrow = UBound(A, 1) - LBound(A, 1) + 1
  ncol = UBound(A, 2) - LBound(A, 2) + 1
  If nrow <> ncol Then GoTo e
  Dim Q, R: ReDim Q(1 To nrow, 1 To nrow) As Double
  
  For k = 1 To nrow
    Q(k, k) = 1
  Next k
  itercount = 0
  chg = 1
  While itercount < maxiter And chg > 0.00000000000001
    eigvec = Q
    QR Application.MMult(A, eigvec), Q, R
    chg = 0
    For k = 1 To nrow
      chg = chg + Q(k, nrow) ^ 2
    Next k
    chg = chg / nrow
    itercount = itercount + 1
  Wend
  
  eigval = GetDiagVector(Application.MMult(Application.Transpose(Q), Application.MMult(A, Q)), False)
  ReDim res(-1 To nrow, 1 To nrow)
  For k = 1 To nrow
    res(0, k) = ""
    res(-1, k) = eigval(LBound(eigval) + k - 1)
    For j = 1 To nrow
      res(k, j) = eigvec(k, j)
    Next j
  Next k
  CovEigenDecompQR = res
e:
End Function
