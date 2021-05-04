VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPivotReportGenerator 
   Caption         =   "Pivot Reports"
   ClientHeight    =   8500.001
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
Private vStorageObjPowerPivotStructure As Variant
Private vStorageObjTableReportStructure As Variant
Private vAllPowerPivotCategories As Variant
Private vAllTableCategories As Variant


Private Sub UserForm_Initialize()
'Populate the userform with report categories as well as an "All" category
    
        
    Set vStorageObjPowerPivotStructure = AssignPivotReportStructureStorage(ActiveWorkbook, False)
    Set vStorageObjTableReportStructure = AssignTableReportStorage(ActiveWorkbook, False)

    If Not vStorageObjTableReportStructure Is Nothing Then
        vAllTableCategories = ReadUniqueReportCategories(vStorageObjTableReportStructure)
        Me.obExcelTableOnly.Value = True
    Else
        Me.obExcelTableOnly.Enabled = False
    End If
    
    If Not vStorageObjPowerPivotStructure Is Nothing Then
        vAllPowerPivotCategories = ReadUniqueReportCategories(vStorageObjPowerPivotStructure)
        Me.obPowerPivotSource = True
    Else
        Me.obPowerPivotSource.Enabled = False
    End If

    RefreshCategories
    
End Sub


Private Sub cbCancel_Click()
    bCancelled = True
    Me.Hide
End Sub

Private Sub cbOK_Click()
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
    Set vStorageObjPowerPivotStructure = Nothing
    Set vStorageObjTableReportStructure = Nothing
End Sub


Private Sub RefreshCategories()
    
    Dim Item As Variant
    
    Me.lbCategories.Clear
    Me.lbCategories.AddItem "All"
    
    Select Case True
        Case obPowerPivotSource.Value = True And Not IsNull(vAllPowerPivotCategories)
            For Each Item In vAllPowerPivotCategories
                Me.lbCategories.AddItem Item
            Next Item
        Case obExcelTableOnly.Value = True And Not IsNull(vAllTableCategories)
            For Each Item In vAllTableCategories
                Me.lbCategories.AddItem Item
            Next Item
    End Select

End Sub


Private Sub RefreshReportListBox()
    
    Dim vArrayOfReportNames As Variant
    Dim ReportName As Variant
    
    Me.lbReports.Clear
    
    Select Case True
    
        Case Me.obPowerPivotSource
            If Me.lbCategories = "All" Then
                vArrayOfReportNames = ReadAllReports(vStorageObjPowerPivotStructure)
            Else
                vArrayOfReportNames = ReadPivotReportsByCategory _
                    (vStorageObjPowerPivotStructure, Me.lbCategories.Text)
            End If
            If Not IsNull(vArrayOfReportNames) Then
                For Each ReportName In vArrayOfReportNames
                    Me.lbReports.AddItem ReportName
                Next ReportName
            End If
    
        Case Me.obExcelTableOnly
            If Me.lbCategories = "All" Then
                vArrayOfReportNames = ReadAllReports(vStorageObjTableReportStructure)
            Else
                vArrayOfReportNames = ReadPivotReportsByCategory _
                    (vStorageObjTableReportStructure, Me.lbCategories.Text)
            End If
            If Not IsNull(vArrayOfReportNames) Then
                For Each ReportName In vArrayOfReportNames
                    Me.lbReports.AddItem ReportName
                Next ReportName
            End If
        
    End Select
    

End Sub


