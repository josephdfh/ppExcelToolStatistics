VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtPaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub paste_and_arrange()
    Dim screen As Boolean: screen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Dim R As Range: Set R = shtPaster.Range("Leftto_headersRowStart").Offset(0, 1).Offset(1, 0)
    R.Value = "."
    R.TextToColumns Destination:=R, _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=shtPaster.Range("consecutive_delimitor").Value, _
        Tab:=True, _
        Semicolon:=True, _
        Comma:=True, _
        Space:=True, _
        Other:=False, OtherChar:=";", _
        FieldInfo:=Array(1, 1)
    R.Value = ""
    Set R = shtTemp.Range("A1").Offset(0, 200)
    R.Value = "."
    R.TextToColumns Destination:=R, _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=shtPaster.Range("consecutive_delimitor").Value, _
        Tab:=True, _
        Semicolon:=True, _
        Comma:=True, _
        Space:=True, _
        Other:=False, OtherChar:=";", _
        FieldInfo:=Array(1, 1)
    R.Value = ""
    Set R = shtPaster.Range("Leftto_headersRowStart").Offset(0, 1).Offset(1, 0)
    Dim temploc As Range
    If pasteCurrentClipboardToShtTemp(temploc) = False Then GoTo lbl_exit
    Call pasterClear
    temploc.Copy
    If pasteValue(R) = False Then Exit Sub
    
    
    
    
    Dim Q As Range: Set Q = shtPaster.UsedRange(shtPaster.UsedRange.count).Offset(1, 1)
    Set Q = shtPaster.Range(R, Q)
    Dim p As Range, c As Range
    Dim toDel As Range: Set toDel = shtPaster.Range("Leftto_headersRowStart").Offset(0, 1)
    
    For Each c In Q.Columns
      For Each p In c.Cells
        If Trim(p.Value) <> "" Then GoTo lbl_next_c
      Next p
      
      'reaching here -> the column is empty -> add column to remove list
      If sameRange(toDel, shtPaster.Range("Leftto_headersRowStart").Offset(0, 1)) Then
        Set toDel = c
      Else
        Set toDel = Union(toDel, c)
      End If
lbl_next_c:
    Next c
    toDel.Delete Shift:=xlToLeft
    
    shtPaster.Range(shtPaster.Range("Leftto_headersRowStart").Offset(0, 1), shtPaster.UsedRange(shtPaster.UsedRange.count).Offset(3, 0).End(xlUp)).Select
lbl_exit:
    Application.ScreenUpdating = True
    Application.ScreenUpdating = screen
End Sub

Sub makeTable()
      Call selectDataTable
      shtPaster.ListObjects.Add(xlSrcRange, Selection, , xlYes).name = "DataTable"
End Sub

Sub selectDataTable()
  Dim last As Range: Set last = shtPaster.UsedRange(shtPaster.UsedRange.count)
  If (last.Column >= shtPaster.Range("Leftto_headersRowStart").Offset(0, 1).Column) Then
    shtPaster.Range(shtPaster.Range("Leftto_headersRowStart").Offset(0, 1), last.Offset(3, 0).End(xlUp)).Select
  Else
    shtPaster.Range("Leftto_headersRowStart").Offset(0, 1).Select
  End If
  
End Sub

Sub pasterClear()
  Dim screen As Boolean: screen = Application.ScreenUpdating
  Application.ScreenUpdating = False
  Dim x
  For Each x In shtPaster.ListObjects
    x.Delete
  Next x
  Dim i As Integer
  Dim last As Range
  For i = 1 To 2  'clear 2 times
    Set last = shtPaster.UsedRange(shtPaster.UsedRange.count)
    If last.Column >= shtPaster.Range("Leftto_headersRowStart").Offset(0, 1).Column Then
      shtPaster.Range(shtPaster.Range("Leftto_headersRowStart").Offset(0, 1), last).Select
    Else
      shtPaster.Range("Leftto_headersRowStart").Offset(0, 1).Select
    End If
    Selection.Clear
    Selection.ColumnWidth = 8.43
  Next i
  shtPaster.Range("Leftto_headersRowStart").Offset(0, 1).Select
  Application.ScreenUpdating = True
  Application.ScreenUpdating = screen
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  If sameRange(Target, shtPaster.Range("split")) Then
    Cancel = True
    
      Call paste_and_arrange
      
    Exit Sub
  End If
  
  If sameRange(Target, shtPaster.Range("make_table")) Then
    Cancel = True
    
      Call makeTable
      
    Exit Sub
  End If
  
  If sameRange(Target, shtPaster.Range("clear")) Then
    Cancel = True
    
      Call pasterClear
      
    Exit Sub
  End If
  
  If sameRange(Target, shtPaster.Range("select")) Then
    Cancel = True
    
      Call selectDataTable
      
    Exit Sub
  End If
  
  If sameRange(Target, shtPaster.Range("notepad")) Then
    Cancel = True
    
    Call Shell("notepad.exe", vbNormalFocus)
      
    Exit Sub
  End If
End Sub
