VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
  Call enable
  If Application.Calculation <> xlCalculationAutomatic Then
    Dim rsp
    rsp = MsgBox("Calculation mode is currently not set to automatic. This software has not been tested for mannual calculation mode. So it is highly recommended to turn calculation mode to automatic." & vbNewLine & "Turn calculation mode to automatic?", vbYesNo)
    If rsp = vbYes Then
      Application.Calculation = xlCalculationAutomatic
    Else
      Call MsgBox("The calculation mode remains not automatic. You can turn it on youself at Formulas|Calculation Options|Automatic", vbOKOnly)
    End If
  End If
End Sub
