VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtTsq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  On Error Resume Next
  Dim states(1 To 2) As Boolean
  states(1) = Application.EnableEvents
  states(2) = Application.ScreenUpdating
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  
  If sameRange(Target, shtTsq.Range("clear")) Then
    Cancel = True
      
    Call clear_Tsq
      
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtTsq.Range("Paste_from_Pivot")) Then
    Cancel = True
      
    Call Paste_from_Pivot
      
    GoTo lbl_exit
  End If
  
lbl_exit:
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  On Error GoTo lbl_exit
  Dim states(1 To 3) As Boolean
  states(1) = Application.EnableEvents
  states(2) = Application.ScreenUpdating
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  If Target.Row = shtTsq.Range("Select_y").Row And _
     Target.Column >= shtTsq.Range("DataArea_Tsq[#Headers]").Cells(1, 1).Column And _
     Target.Column < shtTsq.Range("DataArea_Tsq[#Headers]").Cells(1, 1).Column + shtTsq.Range("DataArea_Tsq[#Headers]").count _
  Then
    Call order_y
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtTsq.Range("T_row")) Then
    Call gen_WorkData
    GoTo lbl_exit
  End If

  
lbl_exit:
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub


Sub clear_Tsq()
  On Error Resume Next
  Application.EnableEvents = False
  
  Dim x: For Each x In shtTsq.ListObjects: x.Delete: Next x:   For Each x In shtTsq.ChartObjects:  x.Delete: Next x
  For Each x In Array("n", "p", "T_row", "Tsq")
    shtTsq.Range(CStr(x)).ClearContents
  Next x
  
  Dim R As Range
  Set R = shtTsq.Range("Left2Data").Offset(-1, 1)
  Dim last As Range: Set last = shtTsq.UsedRange(shtTsq.UsedRange.count)
  If last.Column >= R.Column And last.Row >= R.Row Then
    shtTsq.Range(R, last).Clear
    shtTsq.Range(R, last).ColumnWidth = 8.43
  Else
    R.Resize(100, 100).Clear
    R.Resize(100, 100).ColumnWidth = 8.43
  End If
  Set R = shtTsq.Range("Left2Data").Offset(3, 0)
  Set last = shtTsq.UsedRange(shtTsq.UsedRange.count).Offset(1, 1)
  If last.Row >= R.Row Then
    shtTsq.Range(R, last).EntireRow.Delete
  End If
  Set R = shtTsq.Range("AAA1").End(xlToLeft).Offset(0, 5)
  Set last = shtTsq.UsedRange(shtTsq.UsedRange.count)
  If last.Column >= R.Column Then
    shtTsq.Range(R, last).EntireColumn.Delete
  End If
  shtTsq.UsedRange.ClearOutline
  
  
  Application.EnableEvents = True
End Sub


Sub Paste_from_Pivot()
  Call clear_Tsq
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  Dim R As Range:   Set R = shtTsq.Range("Left2Data").Offset(0, 1)
  Dim lst As ListObject: Set lst = shtPivot.ListObjects("DataArea")
  
  If lst.AutoFilter.FilterMode = False Then
    lst.Range.Copy
    R.PasteSpecial Paste:=xlPasteAll
    R.ListObject.name = "DataArea_Tsq"
  Else
    lst.Range.SpecialCells(xlCellTypeVisible).Copy
    R.PasteSpecial Paste:=xlPasteAll
    shtTsq.ListObjects.Add(xlSrcRange, Selection, , xlYes).name = "DataArea_Tsq"
  End If
  R.ListObject.DataBodyRange.Rows.Group
  R.ListObject.HeaderRowRange.EntireColumn.HorizontalAlignment = xlCenter
  
  
  Dim ncol As Integer: ncol = R.ListObject.HeaderRowRange.Columns.count
  shtTsq.Range("n").Value = R.ListObject.DataBodyRange.Rows.count
  
  Dim r2 As Range: Set r2 = R.Offset(R.ListObject.Range.Rows.count + 2, -1)
  Dim i As Integer
  r2.Value = "New Mean"
  For i = 1 To ncol
    r2.Offset(0, i).FormulaLocal = "=" & R.Offset(0, i - 1).AddressLocal
  Next i
  
  
  
  r2.Offset(1, 0).Value = "(H0: mu=)"
  With r2.Offset(1, 1).Resize(1, ncol)
    .Value = 0
    .name = "'" & shtTsq.name & "'!mu0"
    .Style = "Input"
    
    .Font.Bold = True
    .Font.Color = RGB(0, 0, 0)
  End With
  
  
  Set r2 = r2.Offset(3, 0)
  r2.Value = "Transformer"
  With r2.Offset(1, 1).Resize(ncol, ncol)
    .Value = 0
    .name = "'" & shtTsq.name & "'!Transformer"
    .Style = "Input"
    .Font.Bold = True
    .Font.Color = RGB(0, 0, 0)
  End With
  
  
  For i = 1 To ncol
    r2.Offset(0, i).FormulaLocal = "=" & R.Offset(0, i - 1).AddressLocal
    r2.Offset(i, i).Value = 1
  Next i
  
  Application.EnableEvents = True
  Application.ScreenUpdating = True
End Sub


Private Sub order_y()
  Dim p(1 To 5) As Range, R As Range, r2 As Range
  Dim t As ListObject: Set t = shtTsq.ListObjects("DataArea_Tsq")
  Set p(1) = shtTsq.Range("Left2Data").Offset(-1, 1)
  Set p(2) = p(1).Offset(0, t.HeaderRowRange.Columns.count - 1)
  Dim nrow As Long: nrow = t.Range.Rows.count
  Dim tmp As Range: Set tmp = p(2).Offset(0, 500).Resize(nrow + 1, 1)
  For Each R In t.HeaderRowRange.Offset(-1, 0)
    
    If LCase(Trim(R.Value)) <> "y" Then
      For Each r2 In shtTsq.Range(R, p(2))
        If LCase(Trim(r2.Value)) = "y" Then
          'swap r2 with r
          tmp.Value2 = R.Resize(nrow + 1, 1).Value2
          R.Resize(nrow + 1, 1).Value2 = r2.Resize(nrow + 1, 1).Value2
          r2.Resize(nrow + 1, 1).Value2 = tmp.Value2
          GoTo lbl_next_r
        End If
      Next r2
      
    End If
lbl_next_r:
  Next R
  
  Dim ycount As Integer: ycount = 0
  For Each R In t.HeaderRowRange.Offset(-1, 0)
    If LCase(Trim(R.Value)) = "y" Then
      ycount = ycount + 1
    End If
  Next R
  shtTsq.Range("p").Value = ycount
  shtTsq.Range("T_row").Value = ycount
  
  shtTsq.Range("Transformer").Style = "Normal"
  shtTsq.Range("Transformer").Font.Color = RGB(210, 210, 210)
  With shtTsq.Range("Transformer").Cells(1, 1).Resize(shtTsq.Range("T_row").Value, shtTsq.Range("p").Value)
    .Style = "Input"
    .Font.Bold = True
    .Font.Color = RGB(0, 0, 0)
    .name = "'" & shtTsq.name & "'!Transformer"
  End With
  
  
  shtTsq.Range("mu0").Style = "Normal"
  shtTsq.Range("mu0").Font.Color = RGB(210, 210, 210)
  With shtTsq.Range("mu0").Cells(1, 1).Resize(1, shtTsq.Range("p").Value)
    .Style = "Input"
    .Font.Bold = True
    .Font.Color = RGB(0, 0, 0)
    .name = "'" & shtTsq.name & "'!mu0"
  End With
  
  tmp.Clear
  
  Call gen_WorkData
End Sub


Private Sub gen_WorkData()
  Dim p(1 To 8) As Range, R As Range
  Dim t As ListObject: Set t = shtTsq.ListObjects("DataArea_Tsq")
  Set p(1) = t.Range.Cells(1, 1)
  Set p(2) = p(1).Offset(0, t.HeaderRowRange.Columns.count - 1)
  Set p(3) = p(2).Offset(0, 2)
  
  Dim i As Long, j As Long, ycount As Long, Tycount As Long, n As Long
  n = shtTsq.Range("n").Value
  ycount = shtTsq.Range("p").Value
  Tycount = shtTsq.Range("T_row").Value
  p(3).Resize(n * 2 + 1000, 5000).Clear
  
    
  shtTsq.Range("Transformer").Style = "Normal"
  shtTsq.Range("Transformer").Font.Color = RGB(210, 210, 210)
  With shtTsq.Range("Transformer").Cells(1, 1).Resize(shtTsq.Range("T_row").Value, shtTsq.Range("p").Value)
    .Style = "Input"
    .Font.Bold = True
    .Font.Color = RGB(0, 0, 0)
    .name = "'" & shtTsq.name & "'!Transformer"
  End With
  
  shtTsq.Range("mu0").Style = "Normal"
  shtTsq.Range("mu0").Font.Color = RGB(210, 210, 210)
  With shtTsq.Range("mu0").Cells(1, 1).Resize(1, shtTsq.Range("p").Value)
    .Style = "Input"
    .Font.Bold = True
    .Font.Color = RGB(0, 0, 0)
    .name = "'" & shtTsq.name & "'!mu0"
  End With
  
  For j = 1 To Tycount
    p(3).Offset(0, j).Value = "Ty" & j
  Next j
    
  Set p(4) = p(3).Offset(1, 1).Resize(n, Tycount)
  p(4).FormulaArray = "=MMULT(" & p(1).Offset(1, 0).Resize(n, ycount).AddressLocal & " - mu0, TRANSPOSE(Transformer))"
  p(4).name = "'" & shtTsq.name & "'!Ty"
  p(4).EntireColumn.HorizontalAlignment = xlCenter
  
  Set p(5) = p(3).Offset(0, Tycount + 2)
  
  p(5).Value = "T_Mean"
  p(5).Offset(1, 0).Value = "Sample"
  p(5).Offset(2, 0).Value = "Hypothesized"
  With p(5).Offset(2, 1).Resize(1, Tycount)
    .Style = "Input"
    .name = "'" & shtTsq.name & "'!Tmu0"
    .Value = 0
    .Font.Bold = True
    .Font.Color = RGB(0, 0, 0)
  End With
  
  Set p(6) = p(5).Offset(9, 0)
  p(6).Value = "T_Cov"
  For j = 1 To Tycount
    p(5).Offset(0, j).FormulaLocal = "=" & p(3).Offset(0, j).AddressLocal
    p(5).Offset(1, j).FormulaLocal = "=AVERAGE(" & p(3).Offset(1, j).Resize(n, 1).AddressLocal & ")"
    p(6).Offset(0, j).FormulaLocal = "=" & p(3).Offset(0, j).AddressLocal
    p(6).Offset(j, 0).FormulaLocal = "=" & p(3).Offset(0, j).AddressLocal
  Next j
  p(6).Offset(1, 1).Resize(Tycount, Tycount).FormulaArray = "=MCOVARIANCE(Ty)"
  
  'Simultaneous confidence intervals
  Set p(7) = p(5).Offset(3, 0)
  
  'bonferroni simultaneous confidence interval
  p(7).Value = "Bonfer.S.CI Lo"
  p(7).Offset(1, 0).Value = "Bonfer.S.CI Up"
  For i = 1 To Tycount
  '=L5 - TINV(alpha/T_row, n-1) * SQRT(L14/n)
    p(7).Offset(0, i).FormulaLocal = "=" & p(5).Offset(1, i).AddressLocal & " - TINV(alpha/T_row, n-1) * SQRT(" & p(6).Offset(i, i).AddressLocal & "/n)"
    p(7).Offset(1, i).FormulaLocal = "=" & p(5).Offset(1, i).AddressLocal & " + TINV(alpha/T_row, n-1) * SQRT(" & p(6).Offset(i, i).AddressLocal & "/n)"
  Next i
  
  'Scheffe simultaneous confidence interval
  p(7).Offset(2, 0).Value = "Scheff.S.CI Lo"
  p(7).Offset(3, 0).Value = "Scheff.S.CI Up"
  '=$M$5 - SQRT(HOTELLINGTSQ_INV_RT(alpha, T_row, n-1) )* SQRT($M$14/n)
  For i = 1 To Tycount
  '=L5 - TINV(alpha/T_row, n-1) * SQRT(L14/n)
    p(7).Offset(2, i).FormulaLocal = "=" & p(5).Offset(1, i).AddressLocal & " - SQRT(HOTELLINGTSQ_INV_RT(alpha, T_row, n-1)) * SQRT(" & p(6).Offset(i, i).AddressLocal & "/n)"
    p(7).Offset(3, i).FormulaLocal = "=" & p(5).Offset(1, i).AddressLocal & " + SQRT(HOTELLINGTSQ_INV_RT(alpha, T_row, n-1)) * SQRT(" & p(6).Offset(i, i).AddressLocal & "/n)"
  Next i
  
  
  '=n*MMULT(N5:P5-#REF!,MMULT(MINVERSE(#REF!),TRANSPOSE(N5:P5-#REF!)))
  shtTsq.Range("Tsq").FormulaArray = "=n*MMULT(" & p(5).Offset(1, 1).Resize(1, Tycount).AddressLocal & _
                                      " - Tmu0, MMULT(MINVERSE(" & p(6).Offset(1, 1).Resize(Tycount, Tycount).AddressLocal & _
                                      "),TRANSPOSE(" & p(5).Offset(1, 1).Resize(1, Tycount).AddressLocal & " - Tmu0)))"
  
  shtTsq.Range("Tmu0").EntireColumn.HorizontalAlignment = xlCenter
  shtTsq.Range(p(5), p(6).Offset(Tycount + 2, 0)).HorizontalAlignment = xlRight
  shtTsq.ListObjects("DataArea_Tsq").HeaderRowRange.EntireColumn.HorizontalAlignment = xlCenter
  
End Sub
