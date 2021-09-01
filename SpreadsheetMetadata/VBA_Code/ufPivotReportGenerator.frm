VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPivotReportGenerator 
   Caption         =   "Pivot Reports"
   ClientHeight    =   8780.001
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
Private vAllPowerPivotCategories As Variant
Private vAllTableCategories As Variant


Private Sub chkSaveInNewSpreadsheet_Click()
    If Me.chkSaveInNewSpreadsheet.Value = True Then
        Me.chkGenerateIndex.Enabled = True
        Me.ComboReportNumber.Enabled = True
    Else
        Me.chkGenerateIndex.Enabled = False
        Me.ComboReportNumber.Enabled = False

    End If
End Sub


Private Sub UserForm_Initialize()
'Populate the userform with report categories as well as an "All" category

    Dim i As Integer
    Const iThresholdForIndexGeneration As Integer = 5
    Const iMaxThresholdForReportComboBox As Integer = 20


    Set vStorageObjReportStructure = AssignReportStructureStorage(ActiveWorkbook, False)

    Me.obAll.Value = True
    vAllCategories = ReadUniqueSortedReportCategories(vStorageObjReportStructure)
    vAllPowerPivotCategories = ReadUniqueSortedReportCategories _
        (vStorageObjReportStructure, "Pivot")
    vAllTableCategories = ReadUniqueSortedReportCategories _
        (vStorageObjReportStructure, "Table")

    For i = 1 To iMaxThresholdForReportComboBox
        Me.ComboReportNumber.AddItem (i)
    Next i
    Me.ComboReportNumber.ListIndex = iThresholdForIndexGeneration - 1

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
    
        Select Case True
            Case obAll.Value = True
                For Each item In vAllCategories
                    .AddItem item
                Next item
            Case obPowerPivotSource.Value = True
                For Each item In vAllPowerPivotCategories
                    .AddItem item
                Next item
            Case obExcelTableOnly.Value = True
                For Each item In vAllTableCategories
                    .AddItem item
                Next item
        End Select
    
    End With

    Me.lbCategories.ListIndex = 0
    RefreshReportListBox

End Sub


Private Sub RefreshReportListBox()

    Dim vArrayOfReportNames As Variant
    Dim ReportName As Variant
    Dim i As Long

    Me.lbReports.Clear
    Me.lbReports.ColumnCount = 2
   
    
    Select Case True
    
        Case Me.obAll.Value And Me.lbCategories.Text = "All"
            vArrayOfReportNames = ReadReportNames(vStorageObjReportStructure)
            
        Case Me.obAll.Value And Me.lbCategories.Text <> "All"
            vArrayOfReportNames = ReadReportNames(vStorageObjReportStructure, , _
                Me.lbCategories.Text)
            
        Case Me.obPowerPivotSource.Value And Me.lbCategories.Text = "All"
            vArrayOfReportNames = ReadReportNames(vStorageObjReportStructure, "Pivot")
        
        Case Me.obPowerPivotSource.Value And Me.lbCategories.Text <> "All"
            vArrayOfReportNames = ReadReportNames(vStorageObjReportStructure, "Pivot", _
                Me.lbCategories.Text)
        
        Case Me.obExcelTableOnly.Value And Me.lbCategories.Text = "All"
            vArrayOfReportNames = ReadReportNames(vStorageObjReportStructure, "Table")
        
        Case Me.obExcelTableOnly.Value And Me.lbCategories.Text <> "All"
            vArrayOfReportNames = ReadReportNames(vStorageObjReportStructure, "Table", _
                Me.lbCategories.Text)
    
    End Select
    
    
    i = 0
    If Not IsNull(vArrayOfReportNames) Then
        For Each ReportName In vArrayOfReportNames
            Me.lbReports.AddItem
            Me.lbReports.List(i, 0) = ReportName
            Me.lbReports.List(i, 1) = 1
            i = i + 1
        Next ReportName
    End If
    


End Sub


