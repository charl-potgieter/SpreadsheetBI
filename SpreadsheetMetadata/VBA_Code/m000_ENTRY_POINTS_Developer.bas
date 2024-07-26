Attribute VB_Name = "m000_ENTRY_POINTS_Developer"
Option Explicit

Sub GenerateSpreadsheetMetadataActiveWorkbook()

    StandardEntry
    GenerateSpreadsheetMetadata ActiveWorkbook
    MsgBox ("Metadata created")
    StandardExit

End Sub

Sub CreateSpreadsheetFromMetadata()
'Generates spreadsheet from metadata stored in text files in selected folder


    Dim sFolderPath As String
    Dim wkb As Workbook
    Dim fso As FileSystemObject
    Dim InitialSheetOnWorkbookCreation As Worksheet
    Dim StorageListObjFields
    Dim StorageListObjFieldValues
    Dim StorageListObjFieldFormats
    Dim StorageOther
    Dim SheetNames As Variant
    Dim TargetStorageHeaders As Variant
    Dim ColumnValues As Variant
    Dim FileName As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Long
    Dim rng As Range
    Dim ColumnHasFormula As Boolean
    Dim TableValues() As Dictionary
    Dim TargetSheetStorage As ListStorage
    Dim qry As WorkbookQuery
    Dim cn As WorkbookConnection
    Dim LastUsedFolder As String
    Const StorageRefOFLastFolder As String = _
        "Last utilised folder for creating spreadsheet from metadata"

    StandardEntry
    
    'Get folder containing metadata
    LastUsedFolder = GetSundryStorageItem("Last used directory for creating spreadsheet from metadata")
    sFolderPath = GetFolder(LastUsedFolder)
    If sFolderPath = "" Then
        Exit Sub
    End If
    
    'Save the selected folder for future use
    Set fso = New FileSystemObject
    UpdateSundryStorageValueForGivenItem "Last used directory for creating spreadsheet from metadata", _
        fso.GetParentFolderName(sFolderPath)
    ThisWorkbook.Save

    Set wkb = CreateNewWorkbookWithOneSheet
    Set InitialSheetOnWorkbookCreation = wkb.Sheets(1)
    
    'Import VBA code
    ImportVBAModules wkb, sFolderPath & Application.PathSeparator & "VBA_Code"

    'Assign storage for the relevant spreadsheet metadata
    Set StorageListObjFields = CreateListObjFieldStorage( _
        sFolderPath & Application.PathSeparator & "TableStructure" & _
            Application.PathSeparator & "ListObjectFields.txt", _
        wkb)
    Set StorageListObjFieldValues = CreateListObjFieldValuesStorage( _
        sFolderPath & Application.PathSeparator & "TableStructure" & _
            Application.PathSeparator & "ListObjectFieldValues.txt", _
        wkb)
    Set StorageListObjFieldFormats = CreateListObjFieldFormatsStorage( _
        sFolderPath & Application.PathSeparator & "TableStructure" & _
            Application.PathSeparator & "ListObjectFieldFormats.txt", _
        wkb)
    Set StorageOther = CreateOtherStorage( _
        sFolderPath & Application.PathSeparator & "Other" & _
            Application.PathSeparator & "OtherData.txt", _
        wkb)
    
    FileName = GetCreatorFileName(StorageOther)
    FormatCoverSheet InitialSheetOnWorkbookCreation, FileName
    
    If StorageIsEmpty(StorageListObjFields) Then GoTo Exitpoint
    
    'Create table storage and set formulas
    SheetNames = GetSheetNames(StorageListObjFields)
    Set TargetSheetStorage = New ListStorage
    For i = LBound(SheetNames) To UBound(SheetNames)
    
        TargetStorageHeaders = GetListObjHeaders(StorageListObjFields, SheetNames(i))
        TargetSheetStorage.CreateStorage wkb, SheetNames(i), TargetStorageHeaders
    
        'Set Values
        TableValues = GetTableValues(StorageListObjFieldValues, SheetNames(i))
        For k = LBound(TableValues) To UBound(TableValues)
            TargetSheetStorage.InsertFromDictionary TableValues(k)
        Next k
        
        'Ensure at least one row when setting formats and formulas
        If TargetSheetStorage.NumberOfRecords = 0 Then
            TargetSheetStorage.AddBlankRow
        End If
        
        
        For j = LBound(TargetStorageHeaders) To UBound(TargetStorageHeaders)
            
            'Set number format
            TargetSheetStorage.ListObj.ListColumns(TargetStorageHeaders(j)). _
                DataBodyRange.NumberFormat = GetColumnNumberFormat(StorageListObjFieldFormats, _
                SheetNames(i), TargetStorageHeaders(j))
                        
            'Set font colours
            TargetSheetStorage.ListObj.ListColumns(TargetStorageHeaders(j)). _
                DataBodyRange.Font.Color = GetColumnFontColour(StorageListObjFieldFormats, _
                SheetNames(i), TargetStorageHeaders(j))
            
            'Set formulas
            ColumnHasFormula = GetHeaderHasFormula(StorageListObjFields, _
                SheetNames(i), TargetStorageHeaders(j))
            If ColumnHasFormula Then
                TargetSheetStorage.ListObj.ListColumns(TargetStorageHeaders(j)). _
                    DataBodyRange.Formula = GetColumnFormula(StorageListObjFields, _
                    SheetNames(i), TargetStorageHeaders(j))
            End If
            
        Next j
        
        'Do below to ensure values are formatted per cell format
        'Cell format put in place after values to avoid issues with blank cells.
        'A bit messy but seems to be simplest approach
        For Each rng In TargetSheetStorage.ListObj.DataBodyRange
            rng.Formula = rng.Formula
        Next rng
        
    Next i
    

Exitpoint:
    'Cleanup
    DeleteStorage StorageListObjFields
    DeleteStorage StorageListObjFieldValues
    DeleteStorage StorageListObjFieldFormats
    DeleteStorage StorageOther
    For Each qry In wkb.Queries
        qry.Delete
    Next qry
    For Each cn In wkb.Connections
        cn.Delete
    Next cn

    Set TargetSheetStorage = Nothing
        
    wkb.Activate
    ActiveWindow.WindowState = xlMaximized
    wkb.Sheets(1).Select
    MsgBox ("Spreadsheet created")
        
    StandardExit
    

End Sub



Sub ImportVBAModulesFromFolder()

    Dim wkb As Workbook
    Dim ImportDirectory As String

    StandardEntry
    Set wkb = ActiveWorkbook
    
    If wkb.Name = ThisWorkbook.Name Then
        MsgBox ("Cannot perform this action in current workbook")
        GoTo Exitpoint
    End If
    
    ImportDirectory = GetFolder(wkb.Path)
    ImportVBAModules wkb, ImportDirectory

Exitpoint:
    StandardExit
    Set wkb = Nothing

End Sub


Sub CleanCode()
'Cleans code by exporting, deleting and then re-importing
    
    Dim Response As Integer
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim wkb As Workbook
    Dim TempExportPath As String
    Const TempExportSubFolder As String = "__Temp_VBA_ExportFolder"
    
    StandardEntry
    Set wkb = ActiveWorkbook
    If ThisWorkbook.Name = wkb.Name Then
        MsgBox "Cannot perform this action in current workbook"
        GoTo Exitpoint
    End If
    
    TempExportPath = wkb.Path & Application.PathSeparator & TempExportSubFolder
    If FolderExists(TempExportPath) Then
        On Error Resume Next
        Kill TempExportPath & Application.PathSeparator & "*.*"
        On Error GoTo 0
    Else
        CreateFolder TempExportPath
    End If
    
    ExportVBAModules wkb, TempExportPath
    DeleteEntireVbaProject wkb
    ImportVBAModules wkb, TempExportPath
    
Exitpoint:
    StandardExit
    
End Sub




