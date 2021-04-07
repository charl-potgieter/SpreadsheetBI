VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPivotReportGenerator 
   Caption         =   "Pivot Reports"
   ClientHeight    =   8910.001
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
'@Folder "Reporting"
Option Explicit


Public bCancelled As Boolean
Private vStorageObject As Variant
Private vAllCategories As Variant
Private vCategoriesWithReportsWithExcelBacking As Variant


Private Sub UserForm_Initialize()
'Populate the userform with report categories as well as an "All" category
    
    Dim item As Variant
        
    Set vStorageObject = DataPivotReporting.AssignPivotReportStructureStorage(ActiveWorkbook)

    'Populate the unique report categories on the userform
    vAllCategories = DataPivotReporting.ReadUniquePivotReportCategories(vStorageObject)
    
    Me.lbCategories.AddItem "All"
    If Not IsNull(vAllCategories) Then
        For Each item In vAllCategories
            Me.lbCategories.AddItem item
        Next item
    End If
    
End Sub


Private Sub cbCancel_Click()
    bCancelled = True
    Me.Hide
End Sub

Private Sub cbOK_Click()
    Me.Hide
End Sub


Private Sub chkValueCopy_Click()

    If Me.chkValueCopy.Value = False Then
        Me.chkRetainLiveReport.Value = True
        Me.chkRetainLiveReport.Enabled = False
    Else
        Me.chkRetainLiveReport.Enabled = True
    End If

End Sub

Private Sub lbCategories_Click()
    RefreshReportListBox
End Sub

Private Sub obPowePivotSource_Click()

    RefreshReportListBox
    Me.chkValueCopy.Value = False
    Me.chkRetainLiveReport.Value = True
    Me.chkRetainLiveReport.Enabled = True
    Me.chkValueCopy.Enabled = True
    
End Sub


Private Sub obExcelTableSource_Click()

    RefreshReportListBox
    Me.chkValueCopy.Value = True
    Me.chkRetainLiveReport.Value = False
    Me.chkRetainLiveReport.Enabled = False
    Me.chkValueCopy.Enabled = False
    

End Sub


Private Sub obExcelTableOnly_Click()

    RefreshReportListBox
    Me.chkValueCopy.Value = True
    Me.chkRetainLiveReport.Value = False
    Me.chkRetainLiveReport.Enabled = False
    Me.chkValueCopy.Enabled = False

End Sub


Private Sub UserForm_Terminate()
    bCancelled = True
    Set vStorageObject = Nothing
End Sub



Private Sub RefreshReportListBox()
    
    Dim vArrayOfReportNames As Variant
    Dim ReportName As Variant
    Dim bReportsHasExcelTableSource As Boolean

    
    Me.lbReports.Clear
    If Me.lbCategories = "All" Then
        vArrayOfReportNames = DataPivotReporting.ReadAllPivotReports(vStorageObject)
    Else
        vArrayOfReportNames = DataPivotReporting.ReadPivotReportsByCategory(vStorageObject, Me.lbCategories.Text)
    End If
    
    
    If Not IsNull(vArrayOfReportNames) Then
        For Each ReportName In vArrayOfReportNames
            bReportsHasExcelTableSource = DataPivotReporting. _
                ReadPivotReportHasExcelTableSource(vStorageObject, ReportName)
            
            Select Case True
                Case bReportsHasExcelTableSource
                    Me.lbReports.AddItem ReportName
                Case Not bReportsHasExcelTableSource And Me.obPowePivotSource
                    Me.lbReports.AddItem ReportName
            End Select
            
        Next ReportName
    End If

End Sub

