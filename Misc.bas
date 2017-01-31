Attribute VB_Name = "Misc"
Option Explicit

Private Sub exportModules()
  Dim EXPORT_FOLDER As String
  EXPORT_FOLDER = ThisWorkbook.Path
  Dim m As VBComponent
  For Each m In ThisWorkbook.VBProject.VBComponents
    Select Case m.Type
      Case vbext_ct_StdModule: m.Export ThisWorkbook.Path & "\" & m.name & ".bas"
      Case vbext_ct_MSForm: m.Export ThisWorkbook.Path & "\" & m.name & ".frm"
      Case Else: m.Export ThisWorkbook.Path & "\" & m.name & ".cls"
    End Select
  Next m
End Sub

Public Function AddressR1C1(ByRef R As Range) As String
  AddressR1C1 = R.Worksheet.name & "!R" & R.Row & "C" & R.Column
End Function


Public Sub clearShtTemp()
  On Error GoTo lbl_exit
  Dim res As String
  res = MsgBox("Clear temp sheet?", vbYesNo)
  If res = vbYes Then
    shtTemp.UsedRange.Clear
    shtTemp.UsedRange.Clear
  End If
lbl_exit:
End Sub

Public Function existObject(ByVal nameString As String, ByVal objType As String, ByRef ws As Worksheet) As Boolean
  existObject = False
  On Error GoTo lbl_exit
  If objType = "ListObject" Then
    Debug.Print ws.ListObjects(nameString).name
    existObject = True
    Exit Function
  End If
  
  If objType = "ChartObject" Then
    Debug.Print ws.ChartObjects(nameString).name
    existObject = True
    Exit Function
  End If

  If objType = "Range" Then
    Debug.Print ws.Range(nameString).name
    existObject = True
    Exit Function
  End If
lbl_exit:
End Function

Public Function find2(ByVal whatString As String, ByRef afterMe As Range) As Range
  Set find2 = afterMe.Worksheet.Cells.Find(what:=whatString, After:=afterMe, LookIn:= _
            xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext _
            , MatchCase:=True)
End Function


Public Function GetDiagVector(ByVal mat, Optional ByVal intoColumn As Boolean = True)
  If TypeName(mat) = "Range" Then
    mat = mat.Value2
  End If
  If isScalar(mat) Then
    GetDiagVector = mat
    Exit Function
  End If
  If is1D(mat) Then
    GetDiagVector = mat(LBound(mat, 1))
    Exit Function
  End If
  Dim n As Long
  n = Application.Min(UBound(mat, 1) - LBound(mat, 1) + 1, UBound(mat, 2) - LBound(mat, 2) + 1)
  Dim i As Long
  
  Dim res(): ReDim res(1 To n)
  For i = 1 To n
    res(i) = mat(LBound(mat, 1) + i - 1, LBound(mat, 2) + i - 1)
  Next i
  If intoColumn Then
    res = Application.Transpose(res)
  End If
  GetDiagVector = res
End Function

Public Sub highlight_Nonumer_Cells(ByRef mat As Range)
    
    mat.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=NOT(ISNUMBER(" & Replace(mat(1, 1).Address, "$", "") & "))"
    mat.FormatConditions(mat.FormatConditions.count).SetFirstPriority
    With mat.FormatConditions(1).Font
        .Color = -16777024
        .TintAndShade = 0
    End With
    With mat.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599963377788629
    End With
    mat.FormatConditions(1).StopIfTrue = False
End Sub

Public Function is1D(ByRef mat) As Boolean
  is1D = True
  Dim x As Long
  On Error GoTo lbl_exit
  x = UBound(mat, 2)
  is1D = False
lbl_exit:
End Function

Public Function isScalar(ByRef mat) As Boolean
  isScalar = True
  Dim x As Long
  On Error GoTo lbl_exit
  x = UBound(mat, 1)
  isScalar = False
lbl_exit:
End Function

Public Function MakeDiagMatrix(ByVal mat, Optional ByVal byRow As Boolean = True)
  If TypeName(mat) = "Range" Then
    mat = mat.Value2
  End If
  If isScalar(mat) Then
    MakeDiagMatrix = mat
    Exit Function
  End If
  Dim n As Long: n = UBound(mat, 1) - LBound(mat, 1) + 1
  Dim i As Long
  Dim res() As Double
  If is1D(mat) Then

    ReDim res(1 To n, 1 To n) As Double
    For i = 1 To n
      res(i, i) = mat(LBound(mat, 1) + i - 1)
    Next
    MakeDiagMatrix = res
    Exit Function
  End If
  If byRow = False Then
    mat = Application.Transpose(mat)
  End If
  n = UBound(mat, 1) - LBound(mat, 1) + 1
  Dim p As Long: p = UBound(mat, 2) - LBound(mat, 2) + 1
  ReDim res(1 To n * p, 1 To n * p) As Double
  Dim j As Long, k As Long
  k = 1
  For i = LBound(mat, 1) To UBound(mat, 1)
    For j = LBound(mat, 2) To UBound(mat, 2)
      res(k, k) = mat(i, j)
      k = k + 1
    Next j
  Next i
  MakeDiagMatrix = res
End Function


Public Function MCONSTANTDIAG(ByVal n As Long, Optional ByVal val As Double = 1)
On Error GoTo lbl_exit
  Dim res()
  ReDim res(1 To n, 1 To n)
  Dim i As Long
  For i = 1 To n
     res(i, i) = val
  Next i
  MCONSTANTDIAG = res
lbl_exit:
End Function

Public Function MCORRELATION(ByVal mat)
On Error GoTo lbl_exit
'mat in dataframe format, but mat has no column headers
'a column of mat is a variable

  Dim cov: cov = MCOVARIANCE(mat)
  Dim i As Long
  Dim p As Long: p = UBound(cov, 1) - LBound(cov, 1) + 1
  Dim sdInv: ReDim sdInv(1 To p, 1 To p) As Double
  For i = 1 To p
    sdInv(i, i) = 1 / Sqr(cov(i, i))
  Next i
  MCORRELATION = Application.MMult(sdInv, Application.MMult(cov, sdInv))
lbl_exit:
End Function

Public Function MCOVARIANCE(ByVal mat, Optional ByVal Unbiased As Boolean = True)
'mat in dataframe format, but mat has no column headers
'a column of mat is a variable
'Unbiased = true ==> divided by n-1
'Unbiased = false ==> divided by n (MLE)
  If TypeName(mat) = "Range" Then mat = mat.Value2
  Dim n As Long: n = UBound(mat, 1) - LBound(mat, 1) + 1
  Dim p As Long: p = UBound(mat, 2) - LBound(mat, 2) + 1
  Dim i As Long, j As Long
  Dim ones() As Double: ReDim ones(1 To 1, 1 To n) As Double: For i = 1 To n: ones(1, i) = 1: Next i
  Dim colSum: colSum = Application.MMult(ones, mat) 'sums of columns
  
  Dim b: b = Application.MMult(Application.Transpose(colSum), colSum)
  Dim A: A = Application.MMult(Application.Transpose(mat), mat)
  If p = 1 Then
    If Unbiased Then
      MCOVARIANCE = (A(1) - b(1) / n) / (n - 1)
    Else
      MCOVARIANCE = (A(1) - b(1) / n) / n
    End If
    Exit Function
  End If
  Dim res() As Double: ReDim res(1 To p, 1 To p)
  For i = 1 To p
    For j = 1 To p
      If Unbiased Then
        res(i, j) = (CDbl(A(i, j)) - CDbl(b(i, j)) / n) / (n - 1)
      Else
        res(i, j) = (CDbl(A(i, j)) - CDbl(b(i, j)) / n) / n
      End If
    Next j
  Next i
  MCOVARIANCE = res
End Function

Public Function MTr(ByVal mat As Variant) As Variant
On Error GoTo lbl_exit
    If TypeName(mat) = "Range" Then
        mat = mat.Value2
    End If
    If Not IsArray(mat) Then
      MTr = mat
      Exit Function
    End If
    If is1D(mat) Then
      MTr = mat(LBound(mat))
      Exit Function
    End If
    Dim R As Long: R = UBound(mat) - LBound(mat)
    Dim c As Long: c = UBound(mat, 2) - LBound(mat, 2)
    Dim lo As Long, hi As Long
    lo = LBound(mat)
    hi = UBound(mat)
    If R > c Then hi = lo + c
    Dim i As Long
    MTr = 0
    For i = lo To hi Step 1
        MTr = MTr + mat(i, i)
    Next i
    Exit Function
lbl_exit:
    MTr = "#VALUE!"
End Function

Public Function MultipleRegression(ByVal y, ByVal x, Optional ByVal outputRow As Boolean = True)
'b0, b1, b2, ..., b(k-1), rsq, sd0, sd1, sd2,...,sd(k-1),F,dfssr,dfsse
  On Error GoTo lbl_exit
  Dim est: est = Application.LinEst(y, x, True, True)
  Dim k As Long: k = UBound(est, 2)
  Dim res: ReDim res(1 To 1, 1 To k * 2 + 4)
  Dim i As Long
  For i = 1 To k
    res(1, i) = est(1, 1 + k - i)
    res(1, i + k + 1) = est(2, 1 + k - i)
  Next i
  res(1, k + 1) = est(3, 1)
  res(1, k * 2 + 2) = est(4, 1)
  res(1, k * 2 + 4) = est(4, 2)
  res(1, k * 2 + 3) = k - 1
  If outputRow = False Then
    res = Application.Transpose(res)
  End If
  MultipleRegression = res
lbl_exit:
End Function

Public Sub enable()
Attribute enable.VB_ProcData.VB_Invoke_Func = "E\n14"
  Application.EnableEvents = True
  Application.ScreenUpdating = True
End Sub

Public Sub disable()
  Application.EnableEvents = False
  Application.ScreenUpdating = False
End Sub

Public Function MUniform(ByVal nrow As Long, ByVal ncol As Long, ByVal val)
  On Error GoTo lbl_exit
  Dim res()
  ReDim res(1 To nrow, 1 To ncol)
  Dim i As Long, j As Long
  For i = 1 To nrow
    For j = 1 To ncol
      res(i, j) = val
    Next j
  Next i
  MUniform = res
lbl_exit:
End Function

Public Function pasteCurrentClipboardToShtTemp(ByRef loc As Range) As Boolean
  Dim curScr As Boolean: curScr = Application.ScreenUpdating: Application.ScreenUpdating = False
  Dim curEvt As Boolean: curEvt = Application.EnableEvents: Application.EnableEvents = False
  On Error GoTo lbl_exit
  pasteCurrentClipboardToShtTemp = False
  Dim ws As Worksheet: Set ws = ActiveSheet
  shtTemp.Activate
  Dim R As Range: Set R = shtTemp.UsedRange.SpecialCells(xlCellTypeLastCell)
  Set R = R.Offset(2, 0).End(xlToLeft)
  R.Select
  shtTemp.Paste
  Set loc = Selection
  pasteCurrentClipboardToShtTemp = True
  
lbl_exit:
  ws.Activate
  Application.CutCopyMode = False
  Application.ScreenUpdating = curScr
  Application.EnableEvents = curEvt
End Function

Public Function pasteValue(ByRef R As Range) As Boolean
  pasteValue = False
  ThisWorkbook.Activate
  R.Worksheet.Select
  R.Select
  On Error GoTo lbl_paste_xl:
    R.Worksheet.PasteSpecial Format:="Text", Link:=False, DisplayAsIcon:=False
    GoTo lbl_pasted:
lbl_paste_xl:
    On Error GoTo lbl_exit
    R.Worksheet.Paste
lbl_pasted:
    pasteValue = True
lbl_exit:
    Application.CutCopyMode = False
End Function


Public Sub resetTextToColumn()
  Dim R As Range: Set R = shtTemp.Range("A1")
  R.Value = "."
  R.TextToColumns Destination:=R, _
      DataType:=xlDelimited, _
      TextQualifier:=xlDoubleQuote, _
      ConsecutiveDelimiter:=False, _
      Tab:=False, _
      Semicolon:=False, _
      Comma:=False, _
      Space:=False, _
      Other:=False, _
      FieldInfo:=Array(1, 1)
End Sub


Public Function sameRange(ByRef tar As Range, ByRef mar As Range) As Boolean
  sameRange = False
  If tar.Worksheet.name = mar.Worksheet.name And _
     tar.Row = mar.Row And _
     tar.Column = mar.Column _
     Then
     sameRange = True
  End If
End Function


Public Function SUMPOW(ByRef R As Range, ByVal pow As Long)
  SUMPOW = False
  On Error GoTo lbl_exit
  Dim res As Double: res = 0
  Dim x
  For Each x In R.Cells
    res = res + (x.Value ^ pow)
  Next x
  SUMPOW = res
lbl_exit:
End Function

Public Function SUMPOWSHIFT(ByVal mat, ByVal pow As Double, ByVal Shift As Double)

  If TypeName(mat) = "Range" Then mat = mat.Value2
  If is1D(mat) Then
    SUMPOWSHIFT = (mat + Shift) ^ pow
    Exit Function
  End If
  Dim i As Long, j As Long
  For i = LBound(mat, 1) To UBound(mat, 1)
    For j = LBound(mat, 2) To UBound(mat, 2)
      mat(i, j) = (mat(i, j) + Shift) ^ pow
    Next j
  Next i
  SUMPOWSHIFT = Application.Sum(mat)
End Function




Sub RedHighGreenLow(ByRef R As Range)

    R.FormatConditions.AddColorScale ColorScaleType:=3
    R.FormatConditions(R.FormatConditions.count).SetFirstPriority
    R.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
    With R.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With
    R.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValuePercentile
    R.FormatConditions(1).ColorScaleCriteria(2).Value = 50
    With R.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 8711167
        .TintAndShade = 0
    End With
    R.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueHighestValue
    With R.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With
End Sub


Sub xylablel_Scatter(ByRef x As Range, ByRef y As Range, ByRef lbl As Range, ByRef blankCell As Range)
'  On Error Resume Next
' require data is made into a listobject (excel table)
' x, y, and lbl are all header cells
    blankCell.Select
    Dim ws As Worksheet: Set ws = x.Worksheet
    Dim tb As ListObject: Set tb = x.ListObject
    'sort with lbl
    With tb.Sort
        .SortFields.Clear
        .SortFields.Add Key:=lbl, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Dim lblcount As Long
    Dim lblset() As String: ReDim lblset(1 To 100) As String
    Dim i As Long, j As Long, k As Long
    Dim R(1 To 3) As Range
    Dim jump() As Long: ReDim jump(1 To 100) As Long
    Set R(1) = lbl.Offset(2, 0)
    lblcount = 1
    lblset(1) = lbl.Offset(1, 0).Value
    i = 1
    jump(1) = 1
    While R(1).Value <> ""
      i = i + 1
      If R(1).Value <> lblset(lblcount) Then
        lblcount = lblcount + 1
        lblset(lblcount) = R(1).Value
        jump(lblcount) = i
      End If
      Set R(1) = R(1).Offset(1, 0)
    Wend
    jump(lblcount + 1) = i + 1
    ReDim Preserve lblset(1 To lblcount) As String
    ReDim Preserve jump(1 To lblcount + 1) As Long
    
'=== module: setup chart ===
    ws.Shapes.AddChart.Select
    For i = 1 To lblcount
      ActiveChart.SeriesCollection.NewSeries
      ActiveChart.SeriesCollection(i).name = "=""" & lbl.Value & ":" & lblset(i) & """"
      ActiveChart.SeriesCollection(i).XValues = "=" & ws.name & "!" & x.Offset(jump(i)).Resize(jump(i + 1) - jump(i)).AddressLocal
      ActiveChart.SeriesCollection(i).Values = "=" & ws.name & "!" & y.Offset(jump(i)).Resize(jump(i + 1) - jump(i)).AddressLocal
    Next i
    ActiveChart.ChartType = xlXYScatter
    Dim n As Long: n = tb.DataBodyRange.Rows.count
    ActiveChart.Axes(xlCategory).MinimumScale = Application.Min(x.Offset(1, 0).Resize(n, 1))
    ActiveChart.Axes(xlCategory).MaximumScale = Application.Max(x.Offset(1, 0).Resize(n, 1))
    ActiveChart.Axes(xlValue).MinimumScale = Application.Min(y.Offset(1, 0).Resize(n, 1))
    ActiveChart.Axes(xlValue).MaximumScale = Application.Max(y.Offset(1, 0).Resize(n, 1))
    ActiveChart.Axes(xlCategory).HasMinorGridlines = True
    ActiveChart.Axes(xlValue).HasMinorGridlines = True
    Call ActiveChart.SetElement(msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    Call ActiveChart.SetElement(msoElementPrimaryValueAxisTitleRotated)
    ActiveChart.Axes(xlCategory).AxisTitle.Text = x.Value
    ActiveChart.Axes(xlValue).AxisTitle.Text = y.Value
'=== end of module: setup chart ===
End Sub

Sub EllipsePlotCov(cov, Optional mean As Variant = False, Optional ByRef cht As ChartObject = Nothing)
'require: dimension is 2
'not finished
    On Error Resume Next
    If TypeName(mean) = "Boolean" Then
      mean = Array(0, 0)
    End If
    Dim i As Long, j As Long
    Dim obj, mu(1 To 2) As Double
    i = 0
    For Each obj In mean
      i = i + 1
      mu(i) = CDbl(obj)
    Next obj
    
    Dim eigval, eigvec
    CovEigenDecompQR cov, , eigvec, eigval
    
    Dim A As Double: A = Sqr(eigval(1))
    Dim b As Double: b = Sqr(eigval(2))
    
    Dim x(0 To 160) As Double
    Dim y(0 To 160) As Double
    x(0) = -A
    
    For i = 1 To 10
      x(i) = -A + A * i / 100
    Next i
    
    For i = 11 To 19
      x(i) = -A + A / 10 + A * (i - 10) / 50
    Next i
    
    For i = 20 To 39
      x(i) = -A + A * 28 / 100 + A * (i - 19) * 3.6 / 100
    Next i
    
    x(40) = 0
    
    y(0) = 0
    y(40) = b
    For i = 1 To 39
      y(i) = Sqr((1 - (x(i) / A) ^ 2) * b ^ 2)
    Next i
    
    
    For i = 41 To 80
      x(i) = -x(80 - i)
      y(i) = y(80 - i)
    Next i
    
    
    For i = 81 To 160
      x(i) = x(160 - i)
      y(i) = -y(160 - i)
    Next i
    Dim z
    ReDim z(0 To 160, 1 To 2)
    For i = 0 To 160
      z(i, 1) = x(i)
      z(i, 2) = y(i)
    Next i
    
    z = Application.MMult(z, Application.Transpose(eigvec))
    
    For i = 0 To 160
      x(i) = z(i + 1, 1) + mu(1)
      y(i) = z(i + 1, 2) + mu(2)
    Next i
    
    
    If cht Is Nothing Then
      ActiveSheet.Shapes.AddChart.Select
    Else
      
      cht.Activate
    End If
    ActiveChart.ChartType = xlXYScatterSmoothNoMarkers
    ActiveChart.SeriesCollection(1).Delete
    ActiveChart.SeriesCollection.NewSeries
    'ActiveChart.SeriesCollection(1).Name = "a"
    ActiveChart.SeriesCollection(1).XValues = x
    ActiveChart.SeriesCollection(1).Values = y
    ActiveChart.Legend.Delete
    ActiveChart.Axes(xlValue).HasMajorGridlines = True
    ActiveChart.Axes(xlValue).HasMinorGridlines = True
    ActiveChart.Axes(xlCategory).HasMajorGridlines = True
    ActiveChart.Axes(xlCategory).HasMinorGridlines = True
End Sub


Sub NormalQQplot(ByRef x As Range, ByVal varname As String)
  Dim i As Long, n As Long: n = x.count
  Dim R As Range, tmp As Range
  
  Dim Q() As Double:  ReDim Q(1 To n, 1 To 1) As Double
  Dim y() As Double:  ReDim y(1 To n, 1 To 1) As Double
  Set tmp = x.Worksheet.UsedRange(x.Worksheet.UsedRange.count)
  Set tmp = tmp.Offset(0, 10).End(xlUp).Resize(n, 2)
  i = 1
  For Each R In x.Cells
    y(i, 1) = R.Value
    i = i + 1
  Next R
  
  tmp.Columns(1).Value2 = y
  tmp.Columns(1).Worksheet.Sort.SortFields.Clear
  tmp.Columns(1).Worksheet.Sort.SortFields.Add Key:=tmp.Columns(1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With tmp.Columns(1).Worksheet.Sort
      .SetRange tmp.Columns(1)
      .Header = xlNo
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  Dim mean As Double, sd As Double
  mean = Application.Average(x.Value2)
  sd = Application.WorksheetFunction.StDev(x.Value2)
  
  For i = 1 To n
    Q(i, 1) = Application.NormSInv((0.5 + i - 1) / n)
    
    y(i, 1) = (tmp.Cells(i, 1) - mean) / sd
  Next i
  
  
  x.Worksheet.Activate
  x.Columns(1).Select
  ActiveSheet.Shapes.AddChart.Select
  ActiveChart.ChartType = xlXYScatter
  ActiveChart.SeriesCollection(1).XValues = Q
  ActiveChart.SeriesCollection(1).Values = y
  ActiveChart.SeriesCollection.NewSeries
  ActiveChart.SeriesCollection(2).XValues = Q
  ActiveChart.SeriesCollection(2).Values = Q
  ActiveChart.SeriesCollection(2).ChartType = xlXYScatterLinesNoMarkers
  ActiveChart.Legend.Delete
  ActiveChart.SetElement (msoElementPrimaryValueGridLinesMinorMajor)
  ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMinorMajor)
  ActiveChart.SetElement (msoElementChartTitleAboveChart)
  ActiveChart.ChartTitle.Text = "NQQ plot of " & varname
  
  tmp.EntireColumn.Delete
  x.Select
  
End Sub


Sub MahalanobisChisqQQplot(ByRef x As Range, ByVal dataname As String)
'Require: x must be a rectangular consecutive table of real numbers (no headings)
  Dim i As Long, j As Long
  Dim n As Long, p As Long: n = x.Columns(1).SpecialCells(xlCellTypeVisible).Cells.count: p = x.Columns.count
  
  
  Dim mu() As Double: ReDim mu(1 To 1, 1 To p) As Double
  
  Dim y() As Double: ReDim y(1 To n, 1 To 1) As Double
  
  For j = 1 To p
    mu(1, j) = Application.Average(x.Columns(j).SpecialCells(xlCellTypeVisible).Cells.Value2)
  Next j
  Dim t
  Dim covinv: covinv = Application.MInverse(MCOVARIANCE(x.SpecialCells(xlCellTypeVisible).Cells.Value2))
  Dim xrow() As Double: ReDim xrow(1 To 1, 1 To p) As Double
  For i = 1 To n
    For j = 1 To p
      xrow(1, j) = x(i, j) - mu(1, j)
    Next j
    y(i, 1) = Application.MMult(xrow, Application.MMult(covinv, Application.Transpose(xrow)))(1)
  Next i
  
  Dim Q() As Double: ReDim Q(1 To n, 1 To 1) As Double
  For i = 1 To n
    Q(i, 1) = Application.ChiInv(1 - (i - 0.5) / n, p)
  Next i
  
  
  Dim tmp As Range
  Set tmp = x.Worksheet.UsedRange(x.Worksheet.UsedRange.count)
  Set tmp = tmp.Offset(0, 10).End(xlUp).Resize(n, 1)
  
  tmp.Columns(1).Value2 = y
  tmp.Worksheet.Sort.SortFields.Clear
  tmp.Worksheet.Sort.SortFields.Add Key:=tmp.Columns(1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With tmp.Worksheet.Sort
      .SetRange tmp.Columns(1)
      .Header = xlNo
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  x.Worksheet.Activate
  x.Columns(1).Select
  ActiveSheet.Shapes.AddChart.Select
  ActiveChart.ChartType = xlXYScatter
  ActiveChart.SeriesCollection(1).XValues = Q
  ActiveChart.SeriesCollection(1).Values = tmp.Columns(1).Value2
  ActiveChart.SeriesCollection.NewSeries
  ActiveChart.SeriesCollection(2).XValues = Q
  ActiveChart.SeriesCollection(2).Values = Q
  ActiveChart.SeriesCollection(2).ChartType = xlXYScatterLinesNoMarkers
  ActiveChart.Legend.Delete
  ActiveChart.SetElement (msoElementPrimaryValueGridLinesMinorMajor)
  ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMinorMajor)
  ActiveChart.SetElement (msoElementChartTitleAboveChart)
  ActiveChart.ChartTitle.Text = "ChisqQQ plot of Squared Mahalanobis Distance of " & dataname
  
  tmp.EntireColumn.Delete
End Sub

Sub pvalue_Formatting(ByRef Target As Range, Optional ByVal alphaloc As String = "alpha")
    Target.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=" & alphaloc
    Target.FormatConditions(Target.FormatConditions.count).SetFirstPriority
    With Target.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Target.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Target.FormatConditions(1).StopIfTrue = False
    Target.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=" & alphaloc
    Target.FormatConditions(Target.FormatConditions.count).SetFirstPriority
    With Target.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Target.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Target.FormatConditions(1).StopIfTrue = False
    'target.NumberFormat = "0.####"
End Sub

Function CSSP(mat, cIndex, ByRef RangeForSketch As Range) As Variant
'mat is a rectangular table without header
'cIndex is an array of categorial variable indices
  'Returns a Variant(1 to 3)
  
  'CSSP(1) is Within-Group CSSP    (ECSSP or WCSSP)
  'CSSP(2) is Between-Group CSSP   (HCSSP or BCSSP)
  'CSSP(3) is Total CSSP           (TCSSP)
  'CSSP(4) is the group count
  On Error Resume Next
  
  
  Dim p(1 To 4) As Range, R As Range
  Set p(1) = RangeForSketch.Cells(1, 1)
  
  If TypeName(cIndex) = "Range" Then cIndex = cIndex.Value2
  If False = IsArray(cIndex) Then ReDim Preserve cIndex(1 To 1)
  
  If TypeName(mat) = "Range" Then
    Set p(2) = p(1).Resize(mat.Rows.count, mat.Columns.count)
    p(2).Value2 = mat.Value2
  Else
    Set p(2) = p(1).Resize(UBound(mat, 1) - LBound(mat, 1) + 1, UBound(mat, 2) - LBound(mat, 2) + 1)
    p(2).Value2 = mat
  End If
  
  Dim ncol As Long, nrow As Long
  ncol = p(2).Columns.count
  nrow = p(2).Rows.count
  
  
  Dim idx
  
  p(2).Worksheet.Sort.SortFields.Clear
  For Each idx In cIndex
    p(2).Worksheet.Sort.SortFields.Add Key:=p(2).Columns(idx) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  Next idx
  
  
  With p(2).Worksheet.Sort
      .SetRange p(2)
      .Header = xlGuess
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  
  Dim i As Long, j As Long, k As Long, currentFirstRow As Long, jj As Long
  Dim x
  Dim dbl As Double
  Set R = p(2).Cells(1, 1).Resize(1, ncol)
  Dim firstLineOfGroup As Range
  Set firstLineOfGroup = R
  
  Dim yidx() As Long, ny As Long
  ReDim yidx(1 To ncol) As Long
  ny = 0
  For j = 1 To ncol
    For Each x In cIndex
      If j = CLng(x) Then GoTo lbl_901
    Next x
    ny = ny + 1
    yidx(ny) = j
lbl_901:
  Next j
  
  Dim mugrp, y
  ReDim mugrp(1 To nrow, 1 To ny)
  ReDim y(1 To nrow, 1 To ny)
  Dim grpcount As Long
  grpcount = 0
  
  For i = 1 To nrow
    currentFirstRow = firstLineOfGroup.Row - p(2).Cells(1, 1).Row + 1
    For Each idx In cIndex
      If R.Cells(1, CInt(idx)) <> R.Cells(2, CInt(idx)) Then
        'lastLineOfGroup = True
        Set p(3) = Range(firstLineOfGroup, R)
        For j = 1 To ny
          dbl = Application.Average(p(3).Columns(j))
          For k = 1 To p(3).Rows.count
            mugrp(currentFirstRow + k - 1, j) = dbl
            y(currentFirstRow + k - 1, j) = p(3).Cells(k, yidx(j))
          Next k
lbl_next_j:
        Next j
        Set firstLineOfGroup = R.Offset(1, 0) 'for first line of the next group
        grpcount = grpcount + 1
        GoTo lbl_next_i
      End If
    Next idx
lbl_next_i:
    Set R = R.Offset(1, 0)
  Next i
  
  
  Dim res
  ReDim res(1 To 4)
  res(4) = grpcount
  
  Dim var1, var2, muTotal
  ReDim var1(1 To ny, 1 To ny)
  ReDim var2(1 To ny, 1 To ny)
  'E-CSSP
  var1 = Application.MMult(Application.Transpose(y), y)
  var2 = Application.MMult(Application.Transpose(mugrp), mugrp)
  For i = 1 To ny
    For j = 1 To ny
      var1(i, j) = var1(i, j) - var2(i, j)
    Next j
  Next i
  
  res(1) = var1
  
  ReDim muTotal(1 To 1, 1 To ny)
  For j = 1 To ny
    muTotal(1, j) = Application.Average(p(2).Columns(yidx(j)))
  Next j
  var1 = Application.MMult(Application.Transpose(muTotal), muTotal)
  For i = 1 To ny
    For j = 1 To ny
      var2(i, j) = var2(i, j) - var1(i, j) * nrow
    Next j
  Next i
  
  res(2) = var2
  
  var2 = Application.MMult(Application.Transpose(y), y)
  For i = 1 To ny
    For j = 1 To ny
      var2(i, j) = var2(i, j) - var1(i, j) * nrow
    Next j
  Next i
  
  res(3) = var2
  
  CSSP = res
  
End Function

Function LinearCombinationOfStrings(s As Variant, c As Variant) As String
's is an 1D array of strings
'c is an 1D array of doubles for linear combination
'Require lenght of s and c the same
  Dim res As String
  Dim n As Long, i As Long, zero As Long
  Dim x, y
  If TypeName(s) = "Range" Then
    x = s.Value2
    n = s.Cells.count
    ReDim s(1 To n)
    If n = 1 Then
      s(1) = x
    Else
      i = 1
      For Each y In x
        s(i) = y
        i = i + 1
      Next y
    End If
  End If
  
  If TypeName(c) = "Range" Then
    x = c.Value2
    n = c.Cells.count
    ReDim c(1 To n)
    If n = 1 Then
      c(1) = 1
    Else
      i = 1
      For Each y In x
        c(i) = y
        i = i + 1
      Next y
    End If
  End If
  
  If Not IsArray(s) Then
    s = Array(s)
    c = Array(c)
  End If
  
  If UBound(s) - LBound(s) <> UBound(c) - LBound(c) Then
    Exit Function
  End If
  
  
  zero = LBound(s)
  n = UBound(s) - LBound(s) + 1
  For i = 0 To n - 1
    If c(zero + i) <> 0 Then
      If c(zero + i) = 1 Then
        If res = "" Then
          res = s(zero + i)
        Else
          res = res & "+" & s(zero + i)
        End If
      ElseIf c(zero + i) = -1 Then
        res = res & "-" & s(zero + i)
      Else
        If res = "" Then
          res = c(zero + i) & s(zero + i)
        ElseIf c(zero + i) > 0 Then
          res = res & "+" & c(zero + i) & s(zero + i)
        Else
          res = res & c(zero + i) & s(zero + i)
        End If
      End If
    End If
  Next i
  LinearCombinationOfStrings = res
End Function

Public Function rcfind(rowcell As Range, colcell As Range) As Range
    Set rcfind = Intersect(rowcell.EntireRow, colcell.EntireColumn)
End Function

Public Function vlookup2(lookupTable As ListObject, lookupVal, srcHeader As String, desHeader As String) As Range
    Dim tablename As String: tablename = lookupTable.name
    Dim ws As Worksheet: Set ws = lookupTable.Parent
    Dim rowcell As Range
    Set rowcell = ws.Range(tablename & "[" & srcHeader & "]").Find(what:=lookupVal, LookIn:=xlValues, LookAt:=xlWhole)
    Dim colcell As Range
    Set colcell = lookupTable.HeaderRowRange.Find(what:=desHeader, LookIn:=xlValues, LookAt:=xlWhole)
    Set vlookup2 = rcfind(rowcell, colcell)
End Function
