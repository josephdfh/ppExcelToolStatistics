VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtPivot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub pivotRun()
    Dim R As Range
    Dim res As Boolean
    res = pasteCurrentClipboardToShtTemp(R)
    Call pivotClear
    R.Copy
    Dim start As Range: Set start = shtPivot.Range("above_data_start")
    start.Offset(1, 0).Select
    On Error GoTo lbl_paste_xl:
    ActiveSheet.PasteSpecial Format:="Text", Link:=False, DisplayAsIcon:=False
    GoTo lbl_pasted:
lbl_paste_xl:
    On Error GoTo lbl_exit
    start.Offset(1, 0).PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
lbl_pasted:
    Application.CutCopyMode = False
    shtPivot.ListObjects.Add(xlSrcRange, Selection, , xlYes).name = "DataArea"
    shtPivot.Range("n").Formula = "=ROWS(DataArea)"
    shtPivot.Range("p").Formula = "=COLUMNS(DataArea)"
    Dim n As Long, p As Long
    n = CLng(shtPivot.Range("n").Value)
    p = CLng(shtPivot.Range("p").Value)
    
    Dim des As Range: Set des = start.Offset(1, p + 2)
    ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
                                    SourceData:="DataArea", _
                                    Version:=xlPivotTableVersion12).CreatePivotTable _
                                    TableDestination:=AddressR1C1(des), _
                                    tablename:="PivotArea", _
                                    DefaultVersion:=xlPivotTableVersion12
    
    des.Select
lbl_exit:
End Sub

Private Sub getCovariance()
  Application.ScreenUpdating = False
  Dim start As Range: Set start = shtPivot.Range("above_data_start")
  Dim p As Long: p = CLng(shtPivot.Range("p").Value)
  Dim n As Long: n = CLng(shtPivot.Range("n").Value)
  
  
  Dim i As Integer, j As Integer
  Dim xUnion As Range
  Dim headers: ReDim headers(1 To p) As String
  j = 0
  For i = 1 To p
    If LCase(Trim(start.Offset(0, i - 1).Value)) = "x" Then
      j = j + 1
      If j > 1 Then
        Set xUnion = Union(xUnion, start.Offset(0, i - 1))
      ElseIf j = 1 Then
        Set xUnion = start.Offset(0, i - 1)
      End If
      headers(j) = CStr(start.Offset(1, i - 1).Value)
    End If
  Next i
  If j = 0 Then
    MsgBox "Type an x above a column to select it."
    Exit Sub
  End If
  
  ReDim Preserve headers(1 To j) As String
  Dim Data As Range
  Set Data = Intersect(xUnion.EntireColumn, shtPivot.Range("DataArea").SpecialCells(xlCellTypeVisible))
  Dim datacopy As Range: Set datacopy = start.Offset(n + 10000, p + 1)
  Data.Copy
  datacopy.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Dim Q As Range: Set Q = start.Offset(2 * n + p * 100, 0).End(xlUp).Offset(3, 0)
  Q.Value = Data.AddressLocal
  Dim m As Long: m = Data.count / j
  Q.Offset(-1, 0).Value = "nobs=": Q.Offset(-1, 1).Value = m
  For i = 1 To j
    Q.Offset(0, i).Value = headers(i)
    Q.Offset(i, 0).Value = headers(i)
  Next i
  
  
  Q.Offset(1, 1).Resize(j, j).FormulaArray = MCOVARIANCE(datacopy.Resize(m, j).Value2)
  datacopy.Resize(m, j).Clear
  Application.GoTo reference:=Q.Offset(-1, 0), Scroll:=True
  Application.ScreenUpdating = True
  Q.Offset(1, 1).Resize(j, j).Select
  Selection.Copy
End Sub


Sub pivotClear()
  Dim x
  For Each x In shtPivot.ChartObjects: x.Delete: Next x
  For Each x In shtPivot.ListObjects: x.Delete: Next x
  Dim names: names = Array("n", "p")
  For Each x In names
    shtPivot.Range(CStr(x)).Value = ""
  Next x
  Dim p As Range: Set p = shtPivot.Range("above_data_start")
  p.Value = "."
  shtPivot.Range(p.Offset(500000, 0).End(xlUp).Offset(1000, 0), p).EntireRow.Clear
  Dim Q As Range:  Set Q = shtPivot.UsedRange(shtPivot.UsedRange.count)
  If Q.Row >= p.Row Then
    shtPivot.Range(p, Q).Clear
  End If
  shtPivot.Range(p, p.Offset(0, 10000).End(xlToLeft)).ClearContents
  shtPivot.UsedRange.ColumnWidth = 8.43
  Application.GoTo p.Offset(1, 0), True
End Sub

Sub setPivotSummaryByAverage(ByVal whichSummary As String)
  Dim codeSummary As Integer: codeSummary = xlAverage
  Select Case LCase(whichSummary)
    Case "sum": codeSummary = xlSum
    Case "count": codeSummary = xlCount
    Case "average": codeSummary = xlAverage
    Case "max": codeSummary = xlMax
    Case "min": codeSummary = xlMin
    Case "product": codeSummary = xlProduct
    Case "count numbers": codeSummary = xlCountNums
    Case "stddev": codeSummary = xlStDev
    Case "stddevp": codeSummary = xlStDevP
    Case "var": codeSummary = xlVar
    Case "varp": codeSummary = xlVarP
    Case Else: whichSummary = "Average": codeSummary = xlAverage
  End Select
  Dim R As Range
  Dim s As String
  For Each R In shtPivot.PivotTables("PivotArea").DataLabelRange
    s = R.Value
    s = whichSummary & Mid(s, InStr(s, " "))
    With shtPivot.PivotTables("PivotArea").PivotFields(R.Value)
        .Caption = s
        .Function = codeSummary
    End With
  Next R
End Sub

Private Sub studentize()
  Dim p As Integer: p = shtPivot.Range("p").Value
  Dim n As Long: n = shtPivot.Range("n").Value
  
  Dim R As Range, Q As Range
  Dim sd As Double, mu As Double
  Dim t As ListObject: Set t = shtPivot.ListObjects("DataArea")
  For Each R In t.HeaderRowRange.Cells
    If LCase(Trim(R.Offset(-1, 0).Value)) = "x" Then
      mu = Application.Average(R.Offset(1, 0).Resize(n, 1).Value2)
      sd = Application.StDev(R.Offset(1, 0).Resize(n, 1).Value2)
      For Each Q In R.Offset(1, 0).Resize(n, 1).Cells
        Q.Value = (Q.Value - mu) / sd
      Next Q
    End If
  Next R
End Sub


Private Sub enabler_Click()
  Call enable
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  On Error Resume Next
  Dim states(1 To 2) As Boolean
  states(1) = Application.EnableEvents
  states(2) = Application.ScreenUpdating
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  
  If Target.Row = shtPivot.ListObjects("DataArea").HeaderRowRange.Row - 1 And _
     Target.Column >= shtPivot.ListObjects("DataArea").HeaderRowRange.Cells(1, 1).Column And _
     Target.Column <= shtPivot.ListObjects("DataArea").HeaderRowRange.Cells(1, shtPivot.ListObjects("DataArea").HeaderRowRange.count).Column _
  Then
    Call reorder
    GoTo lbl_exit
  End If


lbl_exit:
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  On Error Resume Next
  Dim states(1 To 2) As Boolean
  states(1) = Application.EnableEvents
  states(2) = Application.ScreenUpdating
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  If sameRange(Target, shtPivot.Range("run")) Then
    Cancel = True
      
    Call pivotRun
      
    GoTo lbl_exit
  End If
  

  
  If sameRange(Target, shtPivot.Range("clear")) Then
    Cancel = True
      
    Call pivotClear
      
    GoTo lbl_exit
  End If
  
  'clear_charts
    If sameRange(Target, shtPivot.Range("clear_charts")) Then
    Cancel = True
      
    Call clear_charts_pivot
      
    GoTo lbl_exit
  End If
  
  
  If sameRange(Target, shtPivot.Range("chg_column_summary_type")) Then
    Cancel = True
    
    Dim response As String: response = "Average"
    response = InputBox(prompt:="Enter summary type for all columns: Sum, Count, Average, Max, Min, Product, Count numbers, stddev, stddevp, var, varp", Default:="Average")
    Call setPivotSummaryByAverage(response)
    
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtPivot.Range("getCovariance")) Then
    Cancel = True
      
    Call getCovariance
      
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtPivot.Range("studentize")) Then
    Cancel = True
      
    Call studentize
      
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtPivot.Range("NormalityPlots")) Then
    Cancel = True
      
    Call NormalityPlots
      
    GoTo lbl_exit
  End If
  
  
  
lbl_exit:
  Application.EnableEvents = states(1)
  Application.ScreenUpdating = states(2)
End Sub


Private Sub NormalityPlots()
  Dim jointnamestr As String: jointnamestr = "("
  Dim R As Range, joint As Range, col As Range
  Dim i As Long
  i = 0
  For Each R In shtPivot.ListObjects("DataArea").HeaderRowRange
    If LCase(Trim(R.Offset(-1, 0).Value)) = "x" Then
      Set col = Intersect(R.EntireColumn.SpecialCells(xlCellTypeVisible), shtPivot.ListObjects("DataArea").DataBodyRange)
      Call NormalQQplot(col, R.Value)
      jointnamestr = jointnamestr & """" & R.Value & ""","
      i = i + 1
    End If
  Next R
  Mid(jointnamestr, Len(jointnamestr), 1) = ")"
  
  If i > 0 Then
    Set joint = shtPivot.ListObjects("DataArea").DataBodyRange.Cells(1, 1).Resize(shtPivot.ListObjects("DataArea").DataBodyRange.Rows.count, i)
    Call MahalanobisChisqQQplot(joint, jointnamestr)
  End If
End Sub

Private Sub clear_charts_pivot()
  Dim x
  For Each x In shtPivot.ChartObjects
    x.Delete
  Next x
End Sub

Private Sub reorder()
  Dim Q As Range, R As Range, r2 As Range, h As Range, tmp As Range, last As Range
  Dim n As Long, p As Long
  Dim str As String
  Dim t As ListObject: Set t = shtPivot.ListObjects("DataArea")
  n = t.DataBodyRange.Rows.count
  p = t.DataBodyRange.Columns.count
  Set h = t.HeaderRowRange.Offset(-1, 0)
  Set Q = h.Cells(1, p)
  Set last = shtPivot.UsedRange(shtPivot.UsedRange.count)
  Set tmp = last.Offset(0, 10).End(xlUp).Offset(h.Row - 2, 0).Resize(n + 2, 1)
  For Each R In h.Cells
    If Trim(LCase(R.Value)) <> "x" Then
      'order
      For Each r2 In shtPivot.Range(R, Q).Cells
        If Trim(LCase(r2.Value)) = "x" Then
          'swap
          str = R.Offset(1, 0).Value
          tmp.Value2 = r2.Resize(n + 2, 1).Value2
          r2.Resize(n + 2, 1).Value2 = R.Resize(n + 2, 1).Value2
          R.Resize(n + 2, 1).Value2 = tmp.Value2
          r2.Offset(1, 0).Value = str
          GoTo lbl_next_r
        End If
      Next r2
    End If
lbl_next_r:
  Next R
  tmp.EntireColumn.Delete
End Sub


