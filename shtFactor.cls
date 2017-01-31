VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtFactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Sub clear_Factor()
  Application.EnableEvents = False
  
  Dim x: For Each x In shtFactor.ListObjects: x.Delete: Next x:   For Each x In shtFactor.ChartObjects:  x.Delete: Next x
  For Each x In Array("n", "p", "m")
    shtFactor.Range(CStr(x)).ClearContents
  Next x
  shtFactor.Range(shtFactor.Range("Run").Offset(1, 0), shtFactor.Range("Run").Offset(10000, 0).End(xlUp).Offset(2, 0)).Clear
  
  Dim R As Range
  Set R = shtFactor.Range("Left2Data").Offset(-1, 1)
  shtFactor.UsedRange.Columns.ClearOutline
  Dim last As Range: Set last = shtFactor.UsedRange(shtFactor.UsedRange.count)
  If last.Column >= R.Column And last.Row >= R.Row Then
    shtFactor.Range(R, last).ColumnWidth = 8.43
    shtFactor.Range(R, last).Clear
  Else
    R.Resize(100, 100).ColumnWidth = 8.43
    R.Resize(100, 100).Clear
  End If
  
  Set R = shtFactor.Range("A1000").End(xlUp).Offset(2, 0)
  Set last = shtFactor.UsedRange(shtFactor.UsedRange.count)
  If last.Row >= R.Row Then
    shtFactor.Range(R, last).EntireRow.Clear
  End If
  
  Set R = shtFactor.Range("AAA1").End(xlToLeft).Offset(0, 5)
  Set last = shtFactor.UsedRange(shtFactor.UsedRange.count)
  If last.Column >= R.Column Then
    shtFactor.Range(R, last).EntireColumn.Clear
  End If
  shtFactor.UsedRange.ClearOutline
  Application.EnableEvents = True
End Sub

Sub Paste_from_Pivot()

  
  Call clear_Factor
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  Dim R As Range:   Set R = shtFactor.Range("Left2Data").Offset(0, 1)
  
  Dim lst As ListObject: Set lst = shtPivot.ListObjects("DataArea")
  
  If lst.AutoFilter.FilterMode = False Then
    lst.Range.Copy
    R.PasteSpecial Paste:=xlPasteAll
    R.ListObject.name = "DataArea_Factor"
  Else
    lst.Range.SpecialCells(xlCellTypeVisible).Copy
    R.PasteSpecial Paste:=xlPasteAll
    shtFactor.ListObjects.Add(xlSrcRange, Selection, , xlYes).name = "DataArea_Factor"
  End If
  R.ListObject.DataBodyRange.Rows.Group
  
  
  Set R = R.End(xlDown).Offset(2, -1)
  R.Value = "Rotation Mat"
  R.Offset(2, 0).Value = "Caution, change the signs of those rows in this rotation matrix corresponding to columns in the loadings matrix that have different signs than SAS's ""factor patterns""."
  
  R.Offset(1, 0).Value = "(Orthonormal!)"
  Dim p As Integer, m As Integer, i As Integer
  p = shtPivot.ListObjects("DataArea").DataBodyRange.Columns.count
  m = Application.Max(1, Int(p / 2))
  shtFactor.Range("p").Value = p
  shtFactor.Range("m").Value = m
  shtFactor.Range("n").Value = shtPivot.ListObjects("DataArea").DataBodyRange.Rows.count
  R.Offset(0, 1).Resize(p, p).Value = 0
  For i = 0 To p - 1
    R.Offset(0, 1).Offset(i, i).Value = 1
  Next i
  Application.EnableEvents = True
  Application.ScreenUpdating = True
End Sub

Sub compute_factor()
  On Error Resume Next
  Dim R(1 To 25) As Range, i As Long, j As Long, k As Long, n As Long, m As Long, xcount As Long, p As Long
  Dim continuous As Boolean: continuous = False
  Set R(25) = shtFactor.Range("Select_with_x")
  
  Set R(23) = shtFactor.ListObjects("DataArea_Factor").HeaderRowRange
  Set R(24) = shtFactor.ListObjects("DataArea_Factor").DataBodyRange
  R(24).ClearFormats
  n = R(24).Rows.count
  p = R(24).Columns.count
  shtFactor.Range("n") = n
  
  
  
  
  Set R(1) = R(25).Offset(0, p + 3)
  
'=== collect "x"s ===
  Dim pos() As Integer: ReDim pos(1 To p) As Integer
  xcount = 0
  For i = 1 To p
    If R(25).Offset(0, i).Value = "x" Then
      xcount = xcount + 1
      pos(xcount) = i
      Call RedHighGreenLow(R(24).Columns(i))
    End If
  Next i
  If pos(xcount) - pos(1) = xcount - 1 Then continuous = True
  If xcount = 0 Then Exit Sub
  shtFactor.Range("p") = xcount
  m = shtFactor.Range("m").Value
  If m <= 0 Or m > xcount Then m = Application.Max(1, Int(xcount / 2))
  shtFactor.Range("m").Value = m
  Dim rot As Range: Set rot = find2("Rotation Mat", R(25)).Offset(0, 1).Resize(m, m)
  Dim last As Range: Set last = shtFactor.UsedRange(shtFactor.UsedRange.count)
  If last.Column >= R(1).Column - 2 And last.Row >= R(1).Row Then
    shtFactor.Range(R(1).Offset(0, -2), last).Clear
  Else
    R(1).Offset(0, -2).Resize(100, 100).Clear
  End If
  
  ReDim Preserve pos(1 To xcount) As Integer
  Dim Data() As Double: ReDim Data(1 To n, 1 To xcount) As Double
  k = 0
  For j = 1 To p
    If R(25).Offset(0, j).Value = "x" Then
      k = k + 1
      For i = 1 To n Step 1
        Data(i, k) = R(24).Item(i, j).Value
      Next i
    End If
  Next j
  
  
  Dim mean, sd
  Dim data2() As Double: ReDim data2(1 To n, 1 To xcount) As Double
  Dim Q(1 To 10) As Range
  R(1).Offset(1, 0).Value = "Mean"
  Set Q(1) = R(1).Offset(1, 1).Resize(1, xcount)
  R(1).Offset(2, 0).Value = "sd"
  Set Q(2) = R(1).Offset(2, 1).Resize(1, xcount)
  For i = 1 To xcount
    R(1).Offset(0, i).Formula = "=" & R(25).Offset(1, pos(i)).AddressLocal
    R(1).Offset(1, i).Formula = "=AVERAGE(" & R(24).Columns(pos(i)).AddressLocal & ")"
    R(1).Offset(2, i).Formula = "=STDEV(" & R(24).Columns(pos(i)).AddressLocal & ")"
  Next i
  mean = R(1).Offset(1, 1).Resize(1, xcount).Value2
  sd = R(1).Offset(2, 1).Resize(1, xcount).Value2
  
'==== "Principle Component Factor Analysis based on Correlation" ===
  Set R(1) = R(1).Offset(4, 0)
  R(1).Value = "Principle Component Factor Analysis based on Correlation"
  
  Set R(2) = R(1).Offset(1, 0)
  R(2).Value = "Correl"
  
  For i = 1 To xcount
    R(2).Offset(0, i).Formula = "=" & R(25).Offset(1, pos(i)).AddressLocal
    R(2).Offset(i, 0).Formula = "=" & R(25).Offset(1, pos(i)).AddressLocal
  Next i
  Set R(3) = R(2).Offset(1, 1).Resize(xcount, xcount)
  R(3).Value2 = MCORRELATION(Data)
  
  Set R(4) = R(2).Offset(xcount + 2, 0)
  R(4).Offset(1, -1).Value = "Eigenvalues"
  Set R(5) = R(4).Offset(1, 0).Resize(2 + xcount, xcount)
  R(5).FormulaArray = "=CovEigenDecompQR(" & R(3).AddressLocal & ")"
  
  Set R(6) = R(4).Offset(2, 0)
  R(6).Offset(1, -1).Value = "Eigenvectors"
  Set R(7) = R(6).Offset(1, 0).Resize(xcount, xcount)
  'R(7).FormulaArray = "=MatEigenvector(" & R(3).AddressLocal & ",TRANSPOSE(" & R(5).AddressLocal & "))"
  
  Set R(8) = R(6).Offset(0, xcount + 1)
  R(8).Value = "Loadings"
  For i = 1 To xcount
    R(8).Offset(i, 0).Formula = "=" & R(25).Offset(1, pos(i)).AddressLocal
  Next i
  For j = 1 To m
    R(8).Offset(0, j).Value = "F" & j
  Next j
  Set R(9) = R(8).Offset(1, 1).Resize(xcount, m)
  R(9).FormulaArray = "=MMULT(MMULT(" & R(6).Offset(1, 0).Resize(xcount, m).AddressLocal & _
                      ",MakeDiagMatrix(SQRT(" & R(4).Offset(1, 0).Resize(1, m).AddressLocal & ")))," & rot.AddressLocal & ")"
  
  Set R(10) = R(6).Offset(xcount + 2, 0)
  
  R(10).Value = "Loadings * Loadings'"
  Set R(11) = R(10).Offset(1, 0).Resize(xcount, xcount)
  R(11).FormulaArray = "=MMULT(" & R(9).AddressLocal & ",TRANSPOSE(" & R(9).AddressLocal & "))"
  
  Set R(12) = R(8).Offset(0, m + 2)
  R(12).Value = "Commun."
  Set R(13) = R(12).Offset(0, 1)
  R(13).Value = "Spec. Var"
  For i = 1 To xcount
    'r(12).Offset(i, -1).Formula = "=" & r(25).Offset(1, pos(i)).AddressLocal
    R(12).Offset(i, 0).Formula = "=SUMSQ(" & R(8).Offset(i, 1).Resize(1, m).AddressLocal & ")"
    R(13).Offset(i, 0).Formula = "=" & R(2).Offset(i, i).AddressLocal & " - " & R(12).Offset(i, 0).AddressLocal
  Next i
  
  Set R(14) = R(10).Offset(0, xcount + 1)
  R(14).Value = "Psi# (=S-LL')"
  Set R(15) = R(14).Offset(1, 0).Resize(xcount, xcount)
  R(15).FormulaArray = "=" & R(3).AddressLocal & " - " & R(11).AddressLocal

  
  Set R(19) = R(13).Offset(0, 2)
  R(19).Value = "Bartlett's Factor Score Coef"
  '=MMULT(MINVERSE(MMULT(TRANSPOSE($AA$65:$AE$74),MMULT(MakeDiagMatrix(1/AN65:AN74),$AA$65:$AE$74))),MMULT(TRANSPOSE($AA$65:$AE$74),MakeDiagMatrix(1/AN65:AN74)))
  R(19).Offset(1, 0).Resize(xcount, m).FormulaArray = "=TRANSPOSE(MMULT(MINVERSE(MMULT(TRANSPOSE(" & R(9).AddressLocal & _
             "),MMULT(MakeDiagMatrix(1/" & R(13).Offset(1, 0).Resize(xcount, 1).AddressLocal & ")," & R(9).AddressLocal & _
             "))),MMULT(TRANSPOSE(" & R(9).AddressLocal & "),MakeDiagMatrix(1/" & R(13).Offset(1, 0).Resize(xcount, 1).AddressLocal & "))))"

  Set R(20) = R(19).Offset(0, m + 1)
  R(20).Value = "Thomson's Factor Score Coef"
  '=MMULT(TRANSPOSE($AA$65:$AE$74),MINVERSE(Z77:AI86 + MakeDiagMatrix(AN65:AN74)))
  R(20).Offset(1, 0).Resize(xcount, m).FormulaArray = "=TRANSPOSE(MMULT(TRANSPOSE(" & R(9).AddressLocal & "),MINVERSE(" & _
                       R(11).AddressLocal & " + MakeDiagMatrix(" & R(13).Offset(1, 0).Resize(xcount, 1).AddressLocal & "))))"
  
  Set R(18) = R(20).Offset(0, m + 1)
  R(18).Value = "SAS's Std Scoring Coef"
  R(18).Offset(1, 0).Resize(xcount, m).FormulaArray = _
                    "=TRANSPOSE(MMULT(TRANSPOSE(" & R(9).AddressLocal & "),MINVERSE(" & R(3).AddressLocal & ")))"
  

  

  Set R(21) = Intersect(shtFactor.Range("Select_with_x").EntireRow, R(18).Offset(0, m + 2).EntireColumn)

  R(21).Value = "Correl: Bartlett's Factor Scores"
  For j = 1 To m
    R(21).Offset(1, j - 1).Value = "F" & j
  Next j
  If continuous = True Then
    '=MMULT((B5:G224-K4:P4)/K5:P5, TRANSPOSE(AO20:AT22))
    R(21).Offset(2, 0).Resize(n, m).FormulaArray = "=MMULT((" & _
                        shtFactor.Range(R(23).Cells(2, pos(1)), R(23).Cells(1, pos(xcount)).End(xlDown)).AddressLocal & _
                        " - " & Q(1).AddressLocal & ") / " & Q(2).AddressLocal & ", " & _
                        R(19).Offset(1, 0).Resize(xcount, m).AddressLocal & ")"
  Else
    For i = 1 To n
      For j = 1 To xcount
        data2(i, j) = (Data(i, j) - mean(1, j)) / sd(1, j)
      Next j
    Next i
    R(21).Offset(2, 0).Resize(n, m).Value2 = _
            Application.MMult(data2, R(19).Offset(1, 0).Resize(xcount, m).Value2)
  End If
  
  For j = 1 To m
    Call RedHighGreenLow(R(21).Offset(2, j - 1).Resize(n, 1))
  Next j
  
  R(21).Offset(0, m + 1).Value = "Correl: Thomson's Factor Scores"
  For j = 1 To m
    R(21).Offset(1, m + j).Value = "F" & j
  Next j
  If continuous = True Then
    '=MMULT((B5:G224-K4:P4)/K5:P5, TRANSPOSE(AO20:AT22))
    R(21).Offset(2, m + 1).Resize(n, m).FormulaArray = "=MMULT((" & _
                        shtFactor.Range(R(23).Cells(2, pos(1)), R(23).Cells(1, pos(xcount)).End(xlDown)).AddressLocal & _
                        " - " & Q(1).AddressLocal & ") / " & Q(2).AddressLocal & ", " & _
                        R(20).Offset(1, 0).Resize(xcount, m).AddressLocal & ")"
  Else
    R(21).Offset(0, m + 1).Offset(2, 0).Resize(n, m).Value2 = _
            Application.MMult(data2, R(20).Offset(1, 0).Resize(xcount, m).Value2)
  End If
  
  For j = 1 To m
    Call RedHighGreenLow(R(21).Offset(2, m + j).Resize(n, 1))
  Next j
  
  R(21).Offset(0, 2 * m + 2).Value = "Correl: SAS's Std Factor Scores"
  For j = 1 To m
    R(21).Offset(1, 2 * m + 1 + j).Value = "F" & j
  Next j
  If continuous = True Then
    '=MMULT((B5:G224-K4:P4)/K5:P5, TRANSPOSE(AO20:AT22))
    R(21).Offset(2, 2 * m + 2).Resize(n, m).FormulaArray = "=MMULT((" & _
                        shtFactor.Range(R(23).Cells(2, pos(1)), R(23).Cells(1, pos(xcount)).End(xlDown)).AddressLocal & _
                        " - " & Q(1).AddressLocal & ") / " & Q(2).AddressLocal & ", " & _
                        R(18).Offset(1, 0).Resize(xcount, m).AddressLocal & ")"
  Else
    R(21).Offset(0, 2 * m + 2).Offset(2, 0).Resize(n, m).Value2 = _
            Application.MMult(data2, R(18).Offset(1, 0).Resize(xcount, m).Value2)
  End If
  
  For j = 1 To m
    Call RedHighGreenLow(R(21).Offset(2, 2 * m + 1 + j).Resize(n, 1))
  Next j
  
'==== "Principle Component Factor Analysis based on Covariance" ===
  Set R(1) = R(6).Offset(xcount * 2 + 5)
  R(1).Value = "Principle Component Factor Analysis based on Covariance"
  
  Set R(2) = R(1).Offset(1, 0)
  R(2).Value = "Cov"
  For i = 1 To xcount
    R(2).Offset(0, i).Formula = "=" & R(25).Offset(1, pos(i)).AddressLocal
    R(2).Offset(i, 0).Formula = "=" & R(25).Offset(1, pos(i)).AddressLocal
    
  Next i
  Set R(3) = R(2).Offset(1, 1).Resize(xcount, xcount)
  R(3).Value2 = MCOVARIANCE(Data)
  
  Set R(4) = R(2).Offset(xcount + 2, 0)
  R(4).Offset(1, -1).Value = "Eigenvalues"
  Set R(5) = R(4).Offset(1, 0).Resize(2 + xcount, xcount)
  R(5).FormulaArray = "=CovEigenDecompQR(" & R(3).AddressLocal & ")"
  
  Set R(6) = R(4).Offset(2, 0)
  R(6).Offset(1, -1).Value = "Eigenvectors"
  Set R(7) = R(6).Offset(1, 0).Resize(xcount, xcount)
  'R(7).FormulaArray = "=MatEigenvector(" & R(3).AddressLocal & ",TRANSPOSE(" & R(5).AddressLocal & "))"
  
  Set R(8) = R(6).Offset(0, xcount + 1)
  R(8).Value = "Loadings"
  For i = 1 To xcount
    R(8).Offset(i, 0).Formula = "=" & R(25).Offset(1, pos(i)).AddressLocal
  Next i
  For j = 1 To m
    R(8).Offset(0, j).Value = "F" & j
  Next j
  Set R(9) = R(8).Offset(1, 1).Resize(xcount, m)
  R(9).FormulaArray = "=MMULT(MMULT(" & R(6).Offset(1, 0).Resize(xcount, m).AddressLocal & _
                      ",MakeDiagMatrix(SQRT(" & R(4).Offset(1, 0).Resize(1, m).AddressLocal & ")))," & rot.AddressLocal & ")"
  
  Set R(10) = R(6).Offset(xcount + 2, 0)
  
  R(10).Value = "Loadings * Loadings'"
  Set R(11) = R(10).Offset(1, 0).Resize(xcount, xcount)
  R(11).FormulaArray = "=MMULT(" & R(9).AddressLocal & ",TRANSPOSE(" & R(9).AddressLocal & "))"
  
  Set R(12) = R(8).Offset(0, m + 2)
  R(12).Value = "Commun."
  Set R(13) = R(12).Offset(0, 1)
  R(13).Value = "Spec. Var"
  For i = 1 To xcount
    'r(12).Offset(i, -1).Formula = "=" & r(25).Offset(1, pos(i)).AddressLocal
    R(12).Offset(i, 0).Formula = "=SUMSQ(" & R(8).Offset(i, 1).Resize(1, m).AddressLocal & ")"
    R(13).Offset(i, 0).Formula = "=" & R(2).Offset(i, i).AddressLocal & " - " & R(12).Offset(i, 0).AddressLocal
  Next i
  
  Set R(14) = R(10).Offset(0, xcount + 1)
  R(14).Value = "Psi# (=S-LL')"
  Set R(15) = R(14).Offset(1, 0).Resize(xcount, xcount)
  R(15).FormulaArray = "=" & R(3).AddressLocal & " - " & R(11).AddressLocal
  'r(15).FormulaArray = "=MakeDiagMatrix(" & r(13).Offset(1, 0).Resize(xcount, 1).AddressLocal & ")"
  
  
  
  Set R(16) = R(8).Offset(2 * xcount + 4, 0)
  R(16).Value = "Std Loadings"
  For i = 1 To xcount
    R(16).Offset(i, 0).Formula = "=" & R(25).Offset(1, pos(i)).AddressLocal
  Next i
  For j = 1 To m
    R(16).Offset(0, j).Value = "F" & j & " std"
  Next j
  For i = 1 To xcount
    For j = 1 To m
      R(16).Offset(i, j).Formula = "=" & R(8).Offset(i, j).AddressLocal & " / SQRT(" & R(2).Offset(i, i).AddressLocal & ")"
    Next j
  Next i
  
  Set R(17) = R(16).Offset(0, m + 2)
  R(17).Value = "Std Commum."
  R(17).Offset(0, 1).Value = "Std Spec. Var"
'  For i = 1 To xcount
'    r(17).Offset(i, -1).Formula = "=" & r(25).Offset(1, pos(i)).AddressLocal
'  Next i
  
  For i = 1 To xcount
    R(17).Offset(i, 0).Formula = "=" & R(12).Offset(i, 0).AddressLocal & " / " & R(2).Offset(i, i).AddressLocal
    R(17).Offset(i, 1).Formula = "=" & R(12).Offset(i, 1).AddressLocal & " / " & R(2).Offset(i, i).AddressLocal
  Next i
  
  
  
  Set R(19) = R(13).Offset(0, 2)
  R(19).Value = "Bartlett's Factor Score Coef"
  '=MMULT(MINVERSE(MMULT(TRANSPOSE($AA$65:$AE$74),MMULT(MakeDiagMatrix(1/AN65:AN74),$AA$65:$AE$74))),MMULT(TRANSPOSE($AA$65:$AE$74),MakeDiagMatrix(1/AN65:AN74)))
  R(19).Offset(1, 0).Resize(xcount, m).FormulaArray = "=TRANSPOSE(MMULT(MINVERSE(MMULT(TRANSPOSE(" & R(9).AddressLocal & _
             "),MMULT(MakeDiagMatrix(1/" & R(13).Offset(1, 0).Resize(xcount, 1).AddressLocal & ")," & R(9).AddressLocal & _
             "))),MMULT(TRANSPOSE(" & R(9).AddressLocal & "),MakeDiagMatrix(1/" & R(13).Offset(1, 0).Resize(xcount, 1).AddressLocal & "))))"

  Set R(20) = R(19).Offset(0, m + 1)
  R(20).Value = "Thomson's Factor Score Coef"
  '=MMULT(TRANSPOSE($AA$65:$AE$74),MINVERSE(Z77:AI86 + MakeDiagMatrix(AN65:AN74)))
  R(20).Offset(1, 0).Resize(xcount, m).FormulaArray = "=TRANSPOSE(MMULT(TRANSPOSE(" & R(9).AddressLocal & "),MINVERSE(" & _
                       R(11).AddressLocal & " + MakeDiagMatrix(" & R(13).Offset(1, 0).Resize(xcount, 1).AddressLocal & "))))"
  
  Set R(18) = R(20).Offset(0, m + 1)
  R(18).Value = "SAS's Std Scoring Coef"
  R(18).Offset(1, 0).Resize(xcount, m).FormulaArray = _
                    "=TRANSPOSE(MMULT(TRANSPOSE(" & R(16).Offset(1, 1).Resize(xcount, m).AddressLocal & "),MINVERSE(" & _
                    find2("Principle Component Factor Analysis based on Correlation", R(3).EntireColumn.Cells(1, 1)).Offset(2, 1).Resize(xcount, xcount).AddressLocal & ")))"

                                                      

  Set R(21) = R(21).Offset(0, 3 * m + 5)


  R(21).Value = "Cov: Bartlett's Factor Scores"
  For j = 1 To m
    R(21).Offset(1, j - 1).Value = "F" & j
  Next j
  If continuous = True Then
    '=MMULT((B5:G224-K4:P4)/K5:P5, TRANSPOSE(AO20:AT22))
    R(21).Offset(2, 0).Resize(n, m).FormulaArray = "=MMULT((" & _
                        shtFactor.Range(R(23).Cells(2, pos(1)), R(23).Cells(1, pos(xcount)).End(xlDown)).AddressLocal & _
                        " - " & Q(1).AddressLocal & ") / " & Q(2).AddressLocal & ", " & _
                        R(19).Offset(1, 0).Resize(xcount, m).AddressLocal & ")"
  Else
    R(21).Offset(2, 0).Resize(n, m).Value2 = _
            Application.MMult(data2, R(19).Offset(1, 0).Resize(xcount, m).Value2)
  End If
  
  For j = 1 To m
    Call RedHighGreenLow(R(21).Offset(2, j - 1).Resize(n, 1))
  Next j
  
  R(21).Offset(0, m + 1).Value = "Cov: Thomson's Factor Scores"
  For j = 1 To m
    R(21).Offset(1, m + j).Value = "F" & j
  Next j
  If continuous = True Then
    '=MMULT((B5:G224-K4:P4)/K5:P5, TRANSPOSE(AO20:AT22))
    R(21).Offset(2, m + 1).Resize(n, m).FormulaArray = "=MMULT((" & _
                        shtFactor.Range(R(23).Cells(2, pos(1)), R(23).Cells(1, pos(xcount)).End(xlDown)).AddressLocal & _
                        " - " & Q(1).AddressLocal & ") / " & Q(2).AddressLocal & ", " & _
                        R(20).Offset(1, 0).Resize(xcount, m).AddressLocal & ")"
  Else
    For i = 1 To n
      For j = 1 To xcount
        data2(i, j) = (Data(i, j) - mean(1, j))
      Next j
    Next i
    R(21).Offset(0, m + 1).Offset(2, 0).Resize(n, m).Value2 = _
            Application.MMult(data2, R(20).Offset(1, 0).Resize(xcount, m).Value2)
  End If
  
  For j = 1 To m
    Call RedHighGreenLow(R(21).Offset(2, m + j).Resize(n, 1))
  Next j
  
  R(21).Offset(0, 2 * m + 2).Value = "Cov: SAS's Std Factor Scores"
  For j = 1 To m
    R(21).Offset(1, 2 * m + 1 + j).Value = "F" & j
  Next j
  If continuous = True Then
    '=MMULT((B5:G224-K4:P4)/K5:P5, TRANSPOSE(AO20:AT22))
    R(21).Offset(2, 2 * m + 2).Resize(n, m).FormulaArray = "=MMULT((" & _
                        shtFactor.Range(R(23).Cells(2, pos(1)), R(23).Cells(1, pos(xcount)).End(xlDown)).AddressLocal & _
                        " - " & Q(1).AddressLocal & ") / " & Q(2).AddressLocal & ", " & _
                        R(18).Offset(1, 0).Resize(xcount, m).AddressLocal & ")"
  Else
    R(21).Offset(0, 2 * m + 2).Offset(2, 0).Resize(n, m).Value2 = _
            Application.MMult(data2, R(18).Offset(1, 0).Resize(xcount, m).Value2)
  End If
  
  For j = 1 To m
    Call RedHighGreenLow(R(21).Offset(2, 2 * m + 1 + j).Resize(n, 1))
  Next j
  
  shtFactor.UsedRange.Columns.Ungroup
  shtFactor.UsedRange.Columns.Ungroup
  shtFactor.Range(R(1), R(18).Offset(0, m)).Columns.Group
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  On Error Resume Next
  Application.EnableEvents = False
  Application.ScreenUpdating = False
    
  If sameRange(Target, shtFactor.Range("clear_all")) Then
    Cancel = True
      
    Call clear_Factor
      
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtFactor.Range("Paste_from_Pivot")) Then
    Cancel = True
      
    Call Paste_from_Pivot
      
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtFactor.Range("Run")) Then
    Cancel = True
    
    Call compute_factor
      
    GoTo lbl_exit
  End If
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
  
  If Target.Row = shtFactor.Range("m").Row And Target.Column = shtFactor.Range("m").Column Then

    Call compute_factor
    
    GoTo lbl_exit
  End If
  
  If Target.Row = shtFactor.Range("Left2Data").Row - 1 And _
     Target.Column >= shtFactor.Range("DataArea_Factor[#Headers]").Cells(1, 1).Column And _
     Target.Column <= shtFactor.Range("DataArea_Factor[#Headers]").Cells(1, 1).End(xlToRight).Column _
  Then
  
    Call compute_factor
    
    GoTo lbl_exit
  End If
  
lbl_exit:
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub
