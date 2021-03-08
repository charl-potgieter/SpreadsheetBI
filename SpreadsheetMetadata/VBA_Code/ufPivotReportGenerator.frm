VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPivotReportGenerator 
   Caption         =   "Power Pivot Reports"
   ClientHeight    =   6590
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9990.001
   OleObjectBlob   =   "ufPivotReportGenerator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufPivotReportGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bCancelled As Boolean

Private Sub cbCancel_Click()
    bCancelled = True
    Me.Hide
End Sub

Private Sub cbOK_Click()
    Me.Hide
End Sub

Private Sub lbCategories_Click()

    Dim v As Variant
    Dim item As Variant
    
    Me.lbReports.Clear
    v = m001_DataAccess.PR_GetReportsByCategory(Me.lbCategories.Text)
    For Each item In v
        Me.lbReports.AddItem item
    Next item

End Sub

Private Sub UserForm_Terminate()
    bCancelled = True
End Sub
