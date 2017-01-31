VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
 On Error Resume Next
 
  With shtData.Range("RegisteredList").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Settings!" & shtSettings.Range("tblSettingsData[Datatable name]").AddressLocal
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Private Sub pasteFromClipboard()
  Dim srcOnTemp As Range
  If pasteCurrentClipboardToShtTemp(srcOnTemp) = False Then Exit Sub
  Application.CutCopyMode = False
  Dim nrow As Long: nrow = srcOnTemp.Rows.count
  Dim ncol As Long: ncol = srcOnTemp.Columns.count
  Dim loc As Range: Set loc = shtData.Range("A30000").End(xlDown).Offset(0, 1).End(xlUp).Offset(3, 0)
  loc.Resize(nrow, ncol).Value2 = srcOnTemp.Value2
  loc.Resize(nrow, ncol).Select
  ActiveWindow.ScrollRow = loc.Row - 1
  ActiveWindow.ScrollColumn = loc.Column - 1
  
  If MsgBox("Register this new data table?", vbYesNo, "Register this new data table?") = vbYes Then
    Call registerDatatable(shtData.name & "!" & Selection.AddressLocal)
  End If
End Sub

Private Sub registerDatatable(ByVal tableRgAddrStr As String)
start:
  Dim tablename As String
  tablename = InputBox("Name the selected table as: ", "Give it a name!")
  
  If IsError(Application.VLookup(tablename, shtSettings.Range("Settings!tblSettingsData[Datatable name]"), 1, False)) = False Then
    Select Case MsgBox("There is an existing table named as " & tablename & "!! Do you wish to overwrite it using the new table?" & vbCr & "Click Cancel to choose a different name. Click No to leave the table as unregistered", vbYesNoCancel, "Name conflict!")
      Case vbCancel: GoTo start  'Choose a different Name
      Case vbNo: Exit Sub  'The current table will be left as unregistered
      Case vbYes: GoTo action
    End Select
  End If
  'add new name to sheet settings
  shtSettings.ListObjects("tblSettingsData").Range.Cells(1, 1).End(xlDown).Offset(1, 0).Value = tablename
action:
  vlookup2(shtSettings.ListObjects("tblSettingsData"), tablename, "Datatable name", "Datatable Address").Value = tableRgAddrStr
  With shtData.Range("RegisteredList").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Settings!" & shtSettings.Range("tblSettingsData[Datatable name]").AddressLocal
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub



Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  On Error Resume Next
  Application.EnableEvents = False
  If sameRange(Target, shtData.Range("PasteClipboard")) Then
    Cancel = True
      
    Call pasteFromClipboard
      
    GoTo lbl_exit
  End If
  
  If sameRange(Target, shtData.Range("RegisterDatatable")) Then
    Cancel = True
      
    Call registerDatatable(Selection.Worksheet.name & "!" & Selection.AddressLocal)
      
    GoTo lbl_exit
  End If
  
lbl_exit:
  Application.EnableEvents = True
  'Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  On Error Resume Next
  If sameRange(Target, shtData.Range("RegisteredList")) Then
    shtData.Range(vlookup2(shtSettings.ListObjects("tblSettingsData"), Target.Value, "Datatable name", "Datatable Address").Value).Select
    ActiveWindow.ScrollRow = Selection.Cells(1, 1).Row - 1
    ActiveWindow.ScrollColumn = Selection.Cells(1, 1).Column - 1
  End If
End Sub
