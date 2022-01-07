VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_ReportSheetFormat 
   Caption         =   "Report sheet format"
   ClientHeight    =   5240
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5050
   OleObjectBlob   =   "uf_ReportSheetFormat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_ReportSheetFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Public UserCancelled As Boolean

Private Sub cbCancel_Click()
    UserCancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer _
                                       , CloseMode As Integer)
    
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
        UserCancelled = True
    End If
    
End Sub


Private Sub cbOk_Click()
    Me.Hide
End Sub


Private Sub UserForm_Initialize()
    UserCancelled = False
End Sub
