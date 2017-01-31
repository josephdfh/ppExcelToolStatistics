Attribute VB_Name = "Distributions"
Option Explicit
Const PI As Double = 3.14159265358979
Function dmvn(x, mu, cov) As Variant
  dmvn = "#VALUE!"
  On Error GoTo lbl_exit
  If (TypeName(cov) = "Range") Then cov = cov.Value2  'O(p^2)
  Dim p As Long: p = 1
  p = UBound(cov) - LBound(cov) + 1
  Dim y() As Double, m2() As Double, obj, i As Long
    
  If (TypeName(mu) = "Range") Then
    ReDim m2(1 To p) As Double  'O(1)
    i = 1
    For Each obj In mu   'O(p)
      m2(i) = obj.Value
      i = i + 1
    Next obj
    mu = m2
  End If
    
  ReDim y(1 To p) As Double 'O(1)
  i = 1
  For Each obj In x  'O(p)
    y(i) = obj - mu(i)
    i = i + 1
  Next obj
   
  dmvn = (2 * PI) ^ (-0.5 * p) / Sqr(Application.MDeterm(cov)) * _
        Exp(-0.5 * Application.SumProduct(y, Application.MMult(y, Application.MInverse(cov))))  'O(p^2)
lbl_exit:
End Function

Public Function Tsq2F(ByVal Tsq As Double, ByVal p As Double, ByVal k As Double) As Variant
  Tsq2F = "#VALUE!"
  On Error GoTo lbl_exit
  Tsq2F = Tsq * (k - p + 1) / k / p
lbl_exit:
End Function

Public Function F2Tsq(ByVal F As Double, ByVal df1 As Double, ByVal df2 As Double) As Variant
  F2Tsq = df1 * (df1 + df2 - 1) / df2 * F
End Function

Public Function HOTELLINGTSQ_DIST(ByVal x As Double, ByVal p As Double, ByVal k As Double) As Variant
  HOTELLINGTSQ_DIST = "#VALUE!"
  On Error GoTo lbl_exit
  HOTELLINGTSQ_DIST = 1 - Application.FDist(x * (k - p + 1) / k / p, p, k - p + 1)
lbl_exit:
End Function

Public Function HOTELLINGTSQ_DIST_RT(ByVal x As Double, ByVal p As Double, ByVal k As Double) As Variant
  HOTELLINGTSQ_DIST_RT = "#VALUE!"
  On Error GoTo lbl_exit
  HOTELLINGTSQ_DIST_RT = Application.FDist(x * (k - p + 1) / k / p, p, k - p + 1)
lbl_exit:
End Function

Public Function HOTELLINGTSQ_INV(ByVal proba As Double, ByVal p As Double, ByVal k As Double) As Variant
  HOTELLINGTSQ_INV = "#VALUE!"
  On Error GoTo lbl_exit
  
  HOTELLINGTSQ_INV = Application.FInv(1 - proba, p, k - p + 1) * k * p / (k - p + 1)
lbl_exit:
End Function

Public Function HOTELLINGTSQ_INV_RT(ByVal proba As Double, ByVal p As Double, ByVal k As Double) As Variant
  HOTELLINGTSQ_INV_RT = "#VALUE!"
  On Error GoTo lbl_exit

  HOTELLINGTSQ_INV_RT = Application.FInv(proba, p, k - p + 1) * k * p / (k - p + 1)
lbl_exit:
End Function



Public Function WilksLambda2F(Lambda, p, A, b) As Variant
  WilksLambda2F = "#VALUE!"
  On Error GoTo lbl_exit
  Dim R As Double: R = A - (p - b + 1) / 2
  Dim Q As Double: Q = p * b / 2 - 1
  Dim t As Double: t = Sqr((p * p * b * b - 4) / (p * p + b * b - 5))
  WilksLambda2F = (R * t - Q) / (p * b) * (Lambda ^ (-1 / t) - 1)
lbl_exit:
End Function

Public Function WilksLambdaFromF(F, p, A, b) As Variant
  WilksLambdaFromF = "#VALUE!"
  On Error GoTo lbl_exit
  Dim R As Double: R = A - (p - b + 1) / 2
  Dim Q As Double: Q = p * b / 2 - 1
  Dim t As Double: t = Sqr((p * p * b * b - 4) / (p * p + b * b - 5))
  WilksLambdaFromF = (F * p * b / (R * t - Q) + 1) ^ (-t)
lbl_exit:
End Function

Public Function WilksLambdaCDF(Lambda, p, A, b) As Variant
  WilksLambdaCDF = "#VALUE!"
  On Error GoTo lbl_exit
  Dim F As Double: F = WilksLambda2F(Lambda, p, A, b)
  Dim R As Double: R = A - (p - b + 1) / 2
  Dim Q As Double: Q = p * b / 2 - 1
  Dim t As Double: t = Sqr((p * p * b * b - 4) / (p * p + b * b - 5))
  WilksLambdaCDF = 1 - Application.FDist(F, p * b, R * t - Q)
lbl_exit:
End Function

Public Function WilksLambdaQuantile(proba, p, A, b) As Variant
  WilksLambdaQuantile = "#VALUE!"
  On Error GoTo lbl_exit
  Dim R As Double: R = A - (p - b + 1) / 2
  Dim Q As Double: Q = p * b / 2 - 1
  Dim t As Double: t = Sqr((p * p * b * b - 4) / (p * p + b * b - 5))
  Dim F As Double: F = Application.FInv(1 - proba, p * b, R * t - Q)
  WilksLambdaQuantile = WilksLambdaFromF(F, p, A, b)
lbl_exit:
End Function
