VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtDisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub clear_Disc()
  Application.EnableEvents = False
  
  Dim x: For Each x In shtDisc.ListObjects: x.Delete: Next x:   For Each x In shtDisc.ChartObjects:  x.Delete: Next x
  For Each x In Array("n", "p", "k")
    shtDisc.Range(CStr(x)).ClearContents
  Next x
 
  shtDisc.Range("m").Value = 1
  Dim R As Range
  Set R = shtDisc.Range("Left2Data").Offset(-1, 1)
  shtDisc.UsedRange.Columns.ClearOutline
  Dim last As Range: Set last = shtDisc.UsedRange(shtDisc.UsedRange.count)
  If last.Column >= R.Column And last.Row >= R.Row Then
    shtDisc.Range(R, last).ColumnWidth = 8.43
    shtDisc.Range(R, last).Clear
  Else
    R.Resize(100, 100).ColumnWidth = 8.43
    R.Resize(100, 100).Clear
  End If
  
  Set R = shtDisc.Range("A1000").End(xlUp).Offset(2, 0)
  Set last = shtDisc.UsedRange(shtDisc.UsedRange.count)
  If last.Row >= R.Row Then
    shtDisc.Range(R, last).EntireRow.Clear
  End If
  
  Set R = shtDisc.Range("AAA1").End(xlToLeft).Offset(0, 5)
  Set last = shtDisc.UsedRange(shtDisc.UsedRange.count)
  If last.Column >= R.Column Then
    shtDisc.Range(R, last).EntireColumn.Clear
  End If
  
  
  Application.EnableEvents = True
End Sub

Sub run()
On Error Resume Next
'=== read data ===
  Dim R(1 To 30) As Range, iter As Range, tb As ListObject, i As Long, j As Long, k As Long, xcount As Integer, ncount As Integer
  Dim n() As Long, xpos() As Integer, ypos As Integer
  Dim y() As String
  ReDim y(1 To 1000) As String
  Dim tbh As Range, tbd As Range
  Dim s As String
  Set tb = shtDisc.ListObjects("DataArea_Disc")
  Set tbh = tb.HeaderRowRange
  Set tbd = tb.DataBodyRange
  shtDisc.Range("n").Value = tbd.Rows.count
  xcount = 0
  ReDim xpos(1 To 1000) As Integer
  For i = 1 To tbh.count
    If tbh.Item(1, i).Offset(-1, 0).Value = "y" Then
      ypos = i
    ElseIf tbh.Item(1, i).Offset(-1, 0).Value = "x" Then
      xcount = xcount + 1
      xpos(xcount) = i
    End If
  Next i
  shtDisc.Range("p").Value = xcount

  ReDim Preserve xpos(1 To xcount) As Integer
  With tb.Sort
        .SortFields.Clear
        .SortFields.Add Key:=tbh.Item(1, ypos), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
  End With
  Set iter = tbh.Item(1, ypos).Offset(1, 0)
  ReDim n(1 To 1000) As Long
  ncount = 1
  While iter.Offset(1, 0).Value <> ""
    n(ncount) = n(ncount) + 1
    If iter.Offset(1, 0).Value <> iter.Value Then
      ncount = ncount + 1
      y(ncount - 1) = iter.Value
    End If
    Set iter = iter.Offset(1, 0)
  Wend
  n(ncount) = n(ncount) + 1
  shtDisc.Range("k").Value = ncount
  If shtDisc.Range("m").Value < 1 Or shtDisc.Range("m").Value > Application.Min(xcount, ncount - 1) Then
    shtDisc.Range("m").Value = Application.Min(xcount, ncount - 1)
  End If
  y(ncount) = iter.Value
  ReDim Preserve n(1 To ncount) As Long
  ReDim Preserve y(1 To ncount) As String
  Set R(1) = shtDisc.Range("Left2Data")
  Set R(2) = R(1).Offset(0, 1).End(xlToRight).Offset(0, 2)
  R(2).Offset(-1, -1).Resize(tbd.Rows.count + xcount * 10 + 100, (xcount + 10) * 10).Clear
  R(2).Value = "Overall xbar"
  For i = 1 To xcount
    R(2).Offset(1, i).Formula = "=AVERAGE(" & tbd.Columns(xpos(i)).AddressLocal & ")"
  Next i
  Set R(3) = R(2).Offset(3, 0)
  R(3).Value = "Number of Groups"
  R(3).Offset(0, 2) = ncount
  
  Set R(4) = R(3).Offset(2, 0)
  R(4).Value = "Sizes"
  For i = 1 To ncount
    R(4).Offset(i, 0).Value = n(i)
  Next i
  
  Set R(5) = R(4).Offset(0, 1)
  R(5).Value = "Group Means"
  k = 1
  For i = 1 To ncount
    For j = 1 To xcount
      R(5).Offset(i, j - 1).Formula = "=AVERAGE(" & tbh.Item(1, xpos(j)).Offset(k, 0).Resize(n(i), 1).AddressLocal & ")"
    Next j
    k = n(i) + k
  Next i
  
  Set R(6) = R(2).Offset(1, 1).Resize(1, xcount)
  Set R(7) = R(4).Offset(ncount + 2, 0)
  R(7).Value = "H CSSP"
  
  s = "=x_x_x()"
  Set R(8) = R(7).Offset(0, 1).Resize(xcount, xcount)
  R(8).FormulaArray = s
  For i = 1 To ncount
    Set iter = R(5).Offset(i, 0).Resize(1, xcount)
    '=B98 * MMULT(TRANSPOSE(C98:D98-C92:D92), C98:D98-C92:D92)
    s = R(4).Offset(i, 0).AddressLocal & " * MMULT(TRANSPOSE(" & iter.AddressLocal & " - " & R(6).AddressLocal & "), " & iter.AddressLocal & " - " & _
          R(6).AddressLocal & ")" & " + x_x_x()"
    R(8).Replace "x_x_x()", s, xlPart
    
  Next i
  R(8).Replace " + x_x_x()", ""

  
  Set R(9) = R(7).Offset(xcount + 2, 0)
  R(9).Value = "E CSSP"
  Dim mat
  Dim Emat() As Double
  ReDim Emat(1 To xcount, 1 To xcount) As Double
  Dim xrow() As Double
  ReDim xrow(1 To 1, 1 To xcount) As Double
  Dim m As Long, h As Long
  m = 0
  For k = 1 To ncount
    For i = 1 To n(k)
      For j = 1 To xcount
        xrow(1, j) = tbd.Item(i + m, xpos(j)).Value - R(5).Offset(k, j - 1).Value
      Next j
      mat = Application.MMult(Application.Transpose(xrow), xrow)
      For j = 1 To xcount
        For h = 1 To xcount
          Emat(j, h) = Emat(j, h) + mat(j, h)
        Next h
      Next j
    Next i
    m = m + n(k)
  Next k
  R(9).Offset(0, 1).Resize(xcount, xcount).Value2 = Emat
  
  Set R(10) = R(9).Offset(xcount + 2, 0)
  R(10).Value = "T CSSP"
  Dim Tmat() As Double
  ReDim Tmat(1 To xcount, 1 To xcount) As Double
  For i = 1 To tbd.Rows.count
    For j = 1 To xcount
      xrow(1, j) = tbd.Item(i, xpos(j)).Value - R(2).Offset(1, j).Value
    Next j
    mat = Application.MMult(Application.Transpose(xrow), xrow)
      For j = 1 To xcount
        For h = 1 To xcount
          Tmat(j, h) = Tmat(j, h) + mat(j, h)
        Next h
      Next j
  Next i
  R(10).Offset(0, 1).Resize(xcount, xcount).Value2 = Tmat
  
  Set R(11) = R(10).Offset(xcount + 2, 0)
  R(11).Value = "Group covariances"
  Dim cov() As Range
  ReDim cov(1 To ncount) As Range
  h = 0
  mat = Null
  For k = 1 To ncount
    Set cov(k) = R(11).Offset(1 + (k - 1) * (xcount + 2), 0)
    cov(k).Value = "Cov: " & R(1).Offset(0, ypos).Value & "=" & y(k)
    ReDim mat(1 To n(k), 1 To xcount)
    If k > 1 Then
      h = h + n(k - 1)
    End If
    For i = 1 To n(k)
      For j = 1 To xcount
        mat(i, j) = CDbl(R(1).Offset(h + i, xpos(j)).Value)
      Next j
    Next i
    Set cov(k) = cov(k).Offset(1, 1).Resize(xcount, xcount)
    cov(k).Value2 = MCOVARIANCE(mat)
  Next k
  
  Set R(12) = R(11).Offset(1 + (k - 1) * (xcount + 2), 0)
  R(12).Value = "Pooled Group Covs S1/n1 + S2/n2 + ..."
  s = ""
  For i = 1 To ncount
    s = s & "+" & cov(i).AddressLocal & " / " & R(4).Offset(i, 0).AddressLocal
  Next i
  Mid(s, 1, 1) = "="
  R(12).Offset(1, 1).Resize(xcount, xcount).FormulaArray = s
  
  Set R(13) = R(2).Offset(0, xcount + 2)
  R(13).Value = "(E^-1)H"
  R(13).Offset(1, 0).Resize(xcount, xcount).FormulaArray = _
      "=MMULT(MINVERSE(" & R(9).Offset(0, 1).Resize(xcount, xcount).AddressLocal & "), " & R(7).Offset(0, 1).Resize(xcount, xcount).AddressLocal & ")"
  Set R(14) = R(13).Offset(0, xcount + 1)
  R(14).Value = "Eigenvalues"
  R(14).Offset(1, 0).Resize(xcount, xcount + 2).FormulaArray = "=CovEigenDecompQR(" & R(13).Offset(1, 0).Resize(xcount, xcount).AddressLocal & ",,,,True)"
  m = shtDisc.Range("m").Value
  shtDisc.Range("chisq").FormulaArray = "=(n-1-(p+k)/2)*SUM(LN(1+" & R(14).Offset(m, 0).Resize(xcount - m + 1, 1).AddressLocal & "))"
  '=(n-1-(p+k)/2)*SUM(LN(1+T6:T8))
  
  
  Set R(15) = R(14).Offset(0, 2)
  R(15).Value = "Eigenvectors"

                                              
  Set R(16) = R(13).Offset(xcount + 2, 0)
  R(16).Value = "(T^-1)H"
  R(16).Offset(1, 0).Resize(xcount, xcount).FormulaArray = _
      "=MMULT(MINVERSE(" & R(10).Offset(0, 1).Resize(xcount, xcount).AddressLocal & "), " & R(7).Offset(0, 1).Resize(xcount, xcount).AddressLocal & ")"
  Set R(17) = R(16).Offset(0, xcount + 1)
  R(17).Value = "Eigenvalues"
  R(17).Offset(1, 0).Resize(xcount, xcount + 2).FormulaArray = "=CovEigenDecompQR(" & R(16).Offset(1, 0).Resize(xcount, xcount).AddressLocal & ",,,,True)"
  
  Set R(18) = R(17).Offset(0, 2)
  R(18).Value = "Eigenvectors"
  Set R(19) = R(18).Offset(0, xcount + 1)
  R(19).Value = "Cancorr"
  R(19).Offset(1, 0).Resize(xcount, 1).FormulaArray = "=SQRT(" & R(17).Offset(1, 0).Resize(xcount, 1).AddressLocal & ")"

  Set R(20) = R(15).Offset(0, xcount + 2)
  R(20).Value = "Rescale Eigenvectors (first row settable)"
  For i = 1 To xcount
    R(20).Offset(i, 0).Value = R(1).Offset(0, xpos(i)).Value
  Next i
  R(20).Offset(1, 1).Resize(1, xcount).Value2 = R(18).Offset(1, 0).Resize(1, xcount).Value2
  R(20).Offset(2, 1).Resize(xcount - 1, xcount).FormulaArray = "=" & R(15).Offset(2, 0).Resize(xcount - 1, xcount).AddressLocal & _
                     "/" & R(15).Offset(1, 0).Resize(1, xcount).AddressLocal & "*" & R(20).Offset(1, 1).Resize(1, xcount).AddressLocal
  R(20).Offset(1, 1).Resize(1, xcount).Style = "Input"
  With R(20).Offset(1, 1).Resize(1, xcount).Font
      .Bold = True
      .ThemeColor = xlThemeColorLight1
      .TintAndShade = 0
  End With
  
  Set R(22) = R(20).Offset(1, 1).Resize(xcount, m)
  
'=== New point classification ==='
  Set R(21) = R(16).Offset(xcount + 2, 0)
  R(21).Value = "Predicting the label of a new point:"
  For i = 1 To xcount
    R(21).Offset(1, i - 1).Value = R(1).Offset(0, xpos(i)).Value
  Next i
  R(21).Offset(2, 0).Resize(1, xcount).Style = "Input"
  With R(21).Offset(2, 0).Resize(1, xcount).Font
      .Bold = True
      .ThemeColor = xlThemeColorLight1
      .TintAndShade = 0
  End With
  R(21).Offset(1, xcount).Value = R(1).Offset(0, ypos).Value
  R(21).Offset(2, xcount).Style = "Note"
  
  
  Set R(23) = R(21).Offset(0, xcount + 2)
  R(23).Value = "Sq Canonical Distances"
  R(23).Offset(1, 0).Value = R(1).Offset(0, ypos).Value
  R(23).Offset(1, 1).Value = "DistSq"
  For i = 1 To ncount
    R(23).Offset(i + 1, 0).Value = y(i)
    R(23).Offset(i + 1, 1).FormulaArray = "=SUMSQ(MMULT(" & _
            R(21).Offset(2, 0).Resize(1, xcount).AddressLocal & _
            "-" & R(5).Offset(i, 0).Resize(1, xcount).AddressLocal & _
            "," & R(22).AddressLocal & "))"
    '=SUMSQ(MMULT($J$14:$K$14-G10:H10,$U$5:$V$6))
  Next i
  
  R(21).Offset(2, xcount).Formula = "=INDEX(" & R(23).Offset(2, 0).Resize(ncount, 1).AddressLocal & _
                    ",MATCH(MIN(" & R(23).Offset(2, 1).Resize(ncount, 1).AddressLocal & ")," & _
                    R(23).Offset(2, 1).Resize(ncount, 1).AddressLocal & ",0))"
  '=INDEX(U18:U20,MATCH(MIN(V18:V20),V18:V20))
  
  Set R(24) = R(23).Offset(2, 1).Resize(ncount, 1)
  R(24).FormatConditions.AddColorScale ColorScaleType:=3
  R(24).FormatConditions(R(24).FormatConditions.count).SetFirstPriority
  R(24).FormatConditions(1).ColorScaleCriteria(1).Type = _
      xlConditionValueLowestValue
  With R(24).FormatConditions(1).ColorScaleCriteria(1).FormatColor
      .Color = 8109667
      .TintAndShade = 0
  End With
  R(24).FormatConditions(1).ColorScaleCriteria(2).Type = _
      xlConditionValuePercentile
  R(24).FormatConditions(1).ColorScaleCriteria(2).Value = 50
  With R(24).FormatConditions(1).ColorScaleCriteria(2).FormatColor
      .Color = 8711167
      .TintAndShade = 0
  End With
  R(24).FormatConditions(1).ColorScaleCriteria(3).Type = _
      xlConditionValueHighestValue
  With R(24).FormatConditions(1).ColorScaleCriteria(3).FormatColor
      .Color = 7039480
      .TintAndShade = 0
  End With
  
  R(24).FormatConditions.AddTop10
  R(24).FormatConditions(R(24).FormatConditions.count).SetFirstPriority
  With R(24).FormatConditions(1)
      .TopBottom = xlTop10Bottom
      .Rank = 1
      .Percent = False
  End With
  With R(24).FormatConditions(1).Font
      .Bold = True
      .Italic = False
      .TintAndShade = 0
  End With
  With R(24).FormatConditions(1).Borders(xlLeft)
      .LineStyle = xlDot
      .TintAndShade = 0
      .Weight = xlThin
  End With
  With R(24).FormatConditions(1).Borders(xlRight)
      .LineStyle = xlDot
      .TintAndShade = 0
      .Weight = xlThin
  End With
  With R(24).FormatConditions(1).Borders(xlTop)
      .LineStyle = xlDot
      .TintAndShade = 0
      .Weight = xlThin
  End With
  With R(24).FormatConditions(1).Borders(xlBottom)
      .LineStyle = xlDot
      .TintAndShade = 0
      .Weight = xlThin
  End With
  R(24).FormatConditions(1).StopIfTrue = False
  
'=== Resubstitution ==='
  Set R(25) = R(21).Offset(4, 0)
  R(25).Value = "Resubstitution error counts"
  For i = 1 To ncount
    R(25).Offset(i, 0).Value = y(i)
    R(25).Offset(i, 2).Formula = "=" & R(25).Offset(i, 1).AddressLocal & "/" & R(4).Offset(i, 0).AddressLocal
    R(25).Offset(i, 2).NumberFormat = "0.00%"
  Next i
  shtDisc.Range("DataArea_Disc[Resubstitution]").Clear
End Sub

Sub MakeScatterPlotByLabel()
  On Error GoTo lbl_exit
  
  Dim R(1 To 10) As Range, x As Range, y As Range, lbl As Range
  Set R(1) = shtDisc.Range("Left2Data")
  Set R(2) = shtDisc.ListObjects("DataArea_Disc").HeaderRowRange
  Set R(3) = shtDisc.Range("ZZ1").End(xlToLeft).Offset(0, 4)
  Dim count As Integer: count = 0
  For Each R(4) In R(2).Offset(-1, 0)
    If R(4).Value = "x" Then
      count = count + 1
      If count = 1 Then Set x = R(4).Offset(1, 0)
      If count = 2 Then Set y = R(4).Offset(1, 0)
    ElseIf R(4).Value = "y" Then
      Set lbl = R(4).Offset(1, 0)
    End If
  Next R(4)
  
  
  Call xylablel_Scatter(x, y, lbl, R(3))
  
lbl_exit:
End Sub


Sub Paste_from_Pivot()
  
  Call clear_Disc
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  Dim R As Range:   Set R = shtDisc.Range("Left2Data").Offset(0, 1)
  
  shtPivot.ListObjects("DataArea").Range.Copy
  R.PasteSpecial Paste:=xlPasteAll
  R.ListObject.name = "DataArea_Disc"
  
  shtDisc.Range("n").Value = shtPivot.ListObjects("DataArea").DataBodyRange.Rows.count
  shtDisc.ListObjects("DataArea_Disc").DataBodyRange.Rows.Group
  shtDisc.Range("Left2Data").Offset(0, 1).End(xlToRight).Offset(0, 1).Value = "Resubstitution"
  Application.EnableEvents = True
  Application.ScreenUpdating = True
End Sub


Sub Resubstitution()
  Application.ScreenUpdating = True
  On Error GoTo lbl_exit:
  Dim t As ListObject: Set t = shtDisc.ListObjects("DataArea_Disc")
  Dim th As Range: Set th = t.HeaderRowRange
  Dim td As Range: Set td = t.DataBodyRange
  Dim x() As Range
  ReDim x(1 To 100) As Range
  Dim y As Range
  Dim rs As Range: Set rs = shtDisc.Range("DataArea_Disc[[#Headers],[Resubstitution]]")
  rs.Offset(-1, 0).Value = ""
  Dim errcount As Range
  Dim i As Long, j As Long, k As Long, n As Long, xcount As Integer
  xcount = 0
  For i = 1 To th.count
    If th(1, i).Offset(-1, 0).Value = "x" Then
      xcount = xcount + 1
      Set x(xcount) = th(1, i)
    ElseIf th(1, i).Offset(-1, 0).Value = "y" Then
      Set y = th(1, i)
    End If
  Next i
  ReDim Preserve x(1 To xcount) As Range
  Dim p As Range: Set p = find2("Predicting the label of a new point:", rs.Offset(xcount + 2, 1))
  Set p = p.Offset(2, 0)
  Set errcount = find2("Resubstitution error counts", p)
  Set errcount = shtDisc.Range(errcount.Offset(1, 1), errcount.End(xlDown).Offset(0, 1))
  errcount.Clear
  errcount.Value = 0
  shtDisc.Range("DataArea_Disc[Resubstitution]").Clear
  
  k = 1
  For i = 1 To td.Rows.count
    Set y = y.Offset(1, 0)
    If i >= 2 And y.Value <> y.Offset(-1, 0).Value Then
      k = k + 1
    End If
    Set rs = rs.Offset(1, 0)
    'Set cv = cv.Offset(1, 0)
    For j = 1 To xcount
      Set x(j) = x(j).Offset(1, 0)
      p.Offset(0, j - 1).Value = x(j).Value
    Next j
    p.Offset(0, xcount).Calculate
    rs.Value = p.Offset(0, xcount).Value
    If rs.Value <> y.Value Then
      With rs.Font
        .Color = -16776961
        .TintAndShade = 0
      End With
      errcount.Item(k, 1).Value = errcount.Item(k, 1).Value + 1
    End If
  Next i
lbl_exit:
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error Resume Next
  Application.EnableEvents = False
  Application.ScreenUpdating = False
    
  If sameRange(Target, shtDisc.Range("clear_all")) Then
    Cancel = True
      
    Call clear_Disc
      
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtDisc.Range("run")) Then
    Cancel = True
      
    Call run
      
    GoTo lbl_exit
  End If
 
  
  If sameRange(Target, shtDisc.Range("Paste_from_Pivot")) Then
    Cancel = True
      
    Call Paste_from_Pivot
      
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtDisc.Range("Resubstitution")) Then
    Cancel = True
      
    Call Resubstitution
      
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtDisc.Range("scatter")) Then
    Cancel = True
      
    Call MakeScatterPlotByLabel
      
    GoTo lbl_exit
  End If
  
  
lbl_exit:
  Application.EnableEvents = True
  Application.ScreenUpdating = True
End Sub




Private Sub Worksheet_Change(ByVal Target As Range)
   On Error Resume Next
   If sameRange(Target, shtDisc.Range("m")) Then
      
    Call run
      
    GoTo lbl_exit
  End If
lbl_exit:
End Sub
