VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufReportShtFormat 
   Caption         =   "Report Sheet Format"
   ClientHeight    =   6270
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5060
   OleObjectBlob   =   "ufReportShtFormat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufReportShtFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CancelButtonClicked As Boolean

Public Property Get UserCancelled() As Boolean
    UserCancelled = CancelButtonClicked
End Property


Private Sub cbCancel_Click()
    Me.Hide
    CancelButtonClicked = True
End Sub

Private Sub cbOk_Click()
    Me.Hide
End Sub


Private Sub UserForm_Initialize()
    CancelButtonClicked = False
End Sub
