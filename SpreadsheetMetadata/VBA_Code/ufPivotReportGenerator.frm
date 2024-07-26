VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPivotReportGenerator 
   Caption         =   "Pivot Reports"
   ClientHeight    =   8270.001
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




'@Folder "Storage.Reporting"
Option Explicit


Public bCancelled As Boolean
Private vStorageObjReportStructure As Variant
Private vAllCategories As Variant


Private Sub UserForm_Initialize()
'Populate the userform with report categories as well as an "All" category

    Dim i As Integer
    Const iThresholdForIndexGeneration As Integer = 5
    Const iMaxThresholdForReportComboBox As Integer = 20


    Set vStorageObjReportStructure = AssignReportStructureStorage(ActiveWorkbook, False)

    vAllCategories = ReadUniqueSortedReportCategories(vStorageObjReportStructure)

    RefreshCategories
    RefreshReportListBox

End Sub


Private Sub cbCancel_Click()
    bCancelled = True
    Me.Hide
End Sub

Private Sub cbOk_Click()
    Me.Hide
End Sub


Private Sub lbCategories_Click()
    RefreshReportListBox
End Sub

Private Sub obPowerPivotSource_Click()
    RefreshCategories
End Sub

Private Sub obExcelTableSource_Click()
    RefreshCategories
End Sub

Private Sub obExcelTableOnly_Click()
    RefreshCategories
End Sub


Private Sub UserForm_Terminate()
    bCancelled = True
    Set vStorageObjReportStructure = Nothing
End Sub


Private Sub RefreshCategories()

    Dim item As Variant


    With Me.lbCategories
        .Clear
        .AddItem "All"
        For Each item In vAllCategories
            .AddItem item
        Next item
    End With

    Me.lbCategories.ListIndex = 0
    RefreshReportListBox

End Sub


Private Sub RefreshReportListBox()

    Dim vArrayOfReportNames As Variant
    Dim ReportName As Variant
    Dim i As Long

    Me.lbReports.Clear
   
    
    If Me.lbCategories.Text = "All" Then
        vArrayOfReportNames = ReadReportNames(vStorageObjReportStructure)
    Else
        vArrayOfReportNames = ReadReportNames(vStorageObjReportStructure, _
                Me.lbCategories.Text)
    End If
    
    i = 0
    If Not IsNull(vArrayOfReportNames) Then
        For Each ReportName In vArrayOfReportNames
            Me.lbReports.AddItem
            Me.lbReports.List(i, 0) = ReportName
            i = i + 1
        Next ReportName
    End If
    


End Sub


