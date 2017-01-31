VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub clearTempSheet()
  Dim x
  For Each x In shtTemp.ListObjects
    x.Delete
  Next x
  
  For Each x In shtTemp.ChartObjects
    x.Delete
  Next x
  shtTemp.UsedRange.Clear
  shtTemp.UsedRange.Clear
End Sub

