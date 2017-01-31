Attribute VB_Name = "modEig"
Option Explicit
Function norm(v)
  norm = Sqr(Application.SumSq(v))
End Function
Function QR(A, Optional Q, Optional R)
On Error GoTo e
  If TypeName(A) = "Range" Then A = A.Value2
  Dim nrow&, ncol&, L1norm#, i&, itercount As Integer, eigval#
  nrow = UBound(A) - LBound(A) + 1
  ncol = UBound(A, 1) - LBound(A, 1) + 1
  If nrow < ncol Then
    QR = "#Error: nrow must >= ncol for QR"
    Exit Function
  End If
  ReDim Q(1 To nrow, 1 To ncol) As Double
  For k = 1 To nrow
    Q(k, 1) = A(LBound(A, 1) + k - 1, LBound(A, 2))
  Next k
  ReDim R(1 To ncol, 1 To ncol) As Double
  ReDim QR(1 To (nrow + ncol), 1 To ncol) As Double
  Dim i&, j&, k&, y, ans
  For j = 1 To ncol
    ReDim y(1 To nrow) As Double
    For k = 1 To nrow: y(k) = A(k, j)
    For i = 1 To (j - 1)
      ans = 0
      For k = 1 To nrow
        ans = ans + Q(k, i) * y(k)
      Next k
      R(i, j) = ans: QR(nrow + i, j) = R(i, j)
      For k = 1 To nrow
        y(k) = y(k) - R(i, j) * Q(k, i)
      Next k
    Next i
    R(j, j) = norm(y): QR(nrow + j, j) = R(i, j)
    For k = 1 To nrow
      Q(k, j) = y(k) / R(j, j): QR(k, j) = Q(k, j)
    Next k
  Next j
  
e:
End Function
Function MatEigenvalue_max(A, Optional maxIter As Integer = 20)
On Error GoTo e
  If TypeName(A) = "Range" Then A = A.Value2
  Dim nrow&, ncol&, L2norm#, i&, itercount As Integer, eigval#
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
