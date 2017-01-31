VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cov2x2Picker 
   Caption         =   "Pick a 2 x 2 Covariance"
   ClientHeight    =   1605
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   6228
   OleObjectBlob   =   "Cov2x2Picker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Cov2x2Picker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public covAddr As String
Private Sub cBtnDone_Click()
  Me.covAddr = Me.RefEdit1.Value
  Me.Hide
End Sub
