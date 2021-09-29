Attribute VB_Name = "m500_CORE_Developer"
Option Explicit
Option Private Module

Sub GenerateSpreadsheetMetadata(ByVal wkb As Workbook)

'Generates selected spreadsheet data to allow the spreadsheet to be recreated
'via VBA.
'Aspects covered include:
'   - Sheet names
'   - Sheet category
'   - Sheet heading
'   - Table name
'   - Number of table columns
'   -  Listobject number format
'   -  Listobject font colour

    Dim sMetaDataRootPath As String
    Dim sTableStructurePath As String
    Dim sVbaCodePath As String
    Dim sOtherPath As String

    sMetaDataRootPath = wkb.Path & Application.PathSeparator & "SpreadsheetMetadata"
    sTableStructurePath = sMetaDataRootPath & Application.PathSeparator & "TableStructure"
    sVbaCodePath = sMetaDataRootPath & Application.PathSeparator & "VBA_Code"
    sOtherPath = sMetaDataRootPath & Application.PathSeparator & "Other"
    

    'Create folders for storing metadata
    If Not FolderExists(sMetaDataRootPath) Then CreateFolder sMetaDataRootPath
    If Not FolderExists(sTableStructurePath) Then CreateFolder sTableStructurePath
    If Not FolderExists(sVbaCodePath) Then CreateFolder sVbaCodePath
    If Not FolderExists(sOtherPath) Then CreateFolder sOtherPath

    'Delete any old files in above folders
    On Error Resume Next
    Kill sTableStructurePath & Application.PathSeparator & "*.*"
    Kill sVbaCodePath & Application.PathSeparator & "*.*"
    Kill sOtherPath & Application.PathSeparator & "*.*"
    On Error GoTo 0

    'Generate listobject metadata
    GenerateMetadataFileListObjectFields wkb, _
        sTableStructurePath & Application.PathSeparator & "ListObjectFields.txt"
    GenerateMetadataFileListObjectValues wkb, _
        sTableStructurePath & Application.PathSeparator & "ListObjectFieldValues.txt"
    GenerateMetadataFileListObjectFormat wkb, _
        sTableStructurePath & Application.PathSeparator & "ListObjectFieldFormats.txt"

    'Export VBA code
    ExportVBAModules wkb, sVbaCodePath
    
    'Generate other info
    GenerateMetadataOther wkb, _
        sOtherPath & Application.PathSeparator & "OtherData.txt"
    


End Sub



Sub GenerateMetadataFileListObjectFields(ByRef wkb As Workbook, ByVal sFilePathAndName As String)

    Dim i As Long
    Dim sht As Worksheet
    Dim ListStorage As ListStorage
    Dim StorageAssigned As Boolean
    Dim lo As ListObject
    Dim sRowToWrite As String
    Dim sFolderPath As String
    Dim iFileNo As Integer

    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "SheetName|ListObjectName|ListObjectHeader|IsFormula|Formula";
    
    Set ListStorage = New ListStorage
    
    For Each sht In wkb.Worksheets
        StorageAssigned = ListStorage.AssignStorage(wkb, sht.Name)
        If StorageAssigned Then
            Set lo = ListStorage.ListObj
            If ListStorage.IsEmpty Then ListStorage.AddBlankRow
            For i = 1 To lo.HeaderRowRange.Columns.Count
                sRowToWrite = vbCr & _
                    sht.Name & "|" & _
                    lo.Name & "|" _
                    & lo.HeaderRowRange.Cells(i) & "|" & _
                    lo.ListColumns(i).DataBodyRange.Cells(1).HasFormula & "|"
                If lo.ListColumns(i).DataBodyRange.Cells(1).HasFormula Then
                    sRowToWrite = sRowToWrite & lo.ListColumns(i).DataBodyRange.Cells(1).Formula
                End If
                Print #iFileNo, sRowToWrite;
            Next i
        End If
    Next sht
    
    Close #iFileNo
    Set ListStorage = Nothing

End Sub


Sub GenerateMetadataFileListObjectValues(ByRef wkb As Workbook, ByVal sFilePathAndName As String)

    Dim i As Long
    Dim j As Long
    Dim sht As Worksheet
    Dim lo As ListObject
    Dim ListStorage As ListStorage
    Dim StorageAssigned As Boolean
    Dim sRowToWrite As String
    Dim sFolderPath As String
    Dim iFileNo As Integer
    
    
    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "SheetName|ListObjectName|ListObjectHeader|Value";
    
    Set ListStorage = New ListStorage
    
    'Write value row by row
    For Each sht In wkb.Worksheets
        StorageAssigned = ListStorage.AssignStorage(wkb, sht.Name)
        If StorageAssigned Then
            Set lo = ListStorage.ListObj
            If ListStorage.IsEmpty Then ListStorage.AddBlankRow
            For i = 1 To lo.DataBodyRange.Rows.Count
                For j = 1 To lo.HeaderRowRange.Columns.Count
                    If Not (lo.ListColumns(j).DataBodyRange.Cells(1).HasFormula) Then
                        sRowToWrite = vbCr & _
                            sht.Name & "|" & _
                            lo.Name & "|" & _
                            lo.ListColumns(j).Name & "|" & _
                            lo.ListColumns(j).DataBodyRange.Cells(i).Value
                            Print #iFileNo, sRowToWrite;
                    End If
                Next j
            Next i
        End If
    Next sht
    
    Close #iFileNo
    Set ListStorage = Nothing

End Sub


Sub GenerateMetadataFileListObjectFormat(ByRef wkb As Workbook, ByVal sFilePathAndName As String)

    Dim i As Long
    Dim sht As Worksheet
    Dim lo As ListObject
    Dim ListStorage As ListStorage
    Dim StorageAssigned As Boolean
    Dim sRowToWrite As String
    Dim sFolderPath As String
    Dim iFileNo As Integer

    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "SheetName|ListObjectName|ListObjectHeader|NumberFormat|FontColour";
    
    Set ListStorage = New ListStorage
    
    For Each sht In wkb.Worksheets
        StorageAssigned = ListStorage.AssignStorage(wkb, sht.Name)
        If StorageAssigned Then
            If ListStorage.IsEmpty Then ListStorage.AddBlankRow
            Set lo = ListStorage.ListObj
            For i = 1 To lo.HeaderRowRange.Columns.Count
                sRowToWrite = vbCr & _
                    sht.Name & "|" & _
                    lo.Name & "|" & _
                        lo.HeaderRowRange.Cells(i) & "|" & _
                        lo.ListColumns(i).DataBodyRange.Cells(1).NumberFormat & "|" & _
                        lo.ListColumns(i).DataBodyRange.Cells(1).Font.Color
                Print #iFileNo, sRowToWrite;
            Next i
        End If
                
    Next sht
    Close #iFileNo
    Set ListStorage = Nothing

End Sub


Sub GenerateMetadataOther(ByRef wkb As Workbook, ByVal sFilePathAndName As String)

    Dim iFileNo As Integer
    Dim sRowToWrite As String
    Dim fso As FileSystemObject
    
    Set fso = New FileSystemObject

    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "Item|Value";
            
    sRowToWrite = vbCr & _
        "FileName|" & fso.GetBaseName(wkb.Name)
    Print #iFileNo, sRowToWrite;
                
    Close #iFileNo

    Set fso = Nothing

End Sub


Sub FormatCoverSheet(ByVal sht As Worksheet, ByVal FileName As String)

    With sht
        .Activate
        .Move Before:=sht.Parent.Sheets(1)
        .Name = "Cover"
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Size = 16
        .Range("B2").Value = FileName
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.Zoom = 80
    End With

End Sub


Function CreateNewWorkbookWithOneSheet() As Workbook

    Dim wkb As Workbook
    
    Set wkb = Application.Workbooks.Add
    Do While wkb.Sheets.Count > 1
        wkb.Sheets(1).Delete
    Loop

    Set CreateNewWorkbookWithOneSheet = wkb
    Set wkb = Nothing
    
End Function





Public Sub ExportVBAModules(ByRef wkb As Workbook, ByVal sFolderPath As String, _
    Optional ByVal bDeleteFirstRow = False, Optional ByVal PrefixForExports As String = "")

'------------------------------------------------------------------------------------------------------------------------
'   Code inspired by Ron De Bruin and Chip Pearson:
'   https://www.rondebruin.nl/win/s9/win002.htm
'   http://www.cpearson.com/excel/vbe.aspx
'
'
'   Requires references
'    - Microsoft Visual Basic For Applications Extensibility 5.3
'    - Microsoft Scripting Runtime
'
'   Requires Trust Access to VBA module:
'   In Excel 2007-2013, click the Developer tab and then click the Macro Security item. In that dialog,
'   choose Macro Settings and check the Trust access to the VBA project object model
'
'   Be aware that above may trigger action from antivirus software
'
'------------------------------------------------------------------------------------------------------------------------



'Saves active workbook and exports file to sFolderPath
' *****IMPORTANT NOTE****
' Any existing files will be overwritten
'if bDeleteFirstRow is set as true the first row of the file contining module name is deleted to enable file
'to be simply copied and pasted into VBA IDE
'If PrefixForExports is set only modules with that prefix are exported

    Dim sExportFileName As String
    Dim bExport As Boolean
    Dim sFileName As String
    Dim cmpComponent As VBIDE.VBComponent


    If wkb.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
            "not possible to export the code"
        Exit Sub
    End If
    
    For Each cmpComponent In wkb.VBProject.VBComponents
        
        bExport = True
        sFileName = cmpComponent.Name

        'Set filename
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                sFileName = cmpComponent.Name & ".cls"
            Case vbext_ct_MSForm
                sFileName = cmpComponent.Name & ".frm"
            Case vbext_ct_StdModule
                sFileName = cmpComponent.Name & ".bas"
            Case vbext_ct_Document
                ' This is a worksheet or workbook object - don't export.
                bExport = False
        End Select
        
        If PrefixForExports <> "" And Left(sFileName, Len(PrefixForExports)) <> _
            PrefixForExports Then
                bExport = False
        End If
        
        If bExport Then
            sExportFileName = sFolderPath & Application.PathSeparator & sFileName
            cmpComponent.Export sExportFileName
            If bDeleteFirstRow Then
                DeleteFirstLineOfTextFile sExportFileName
            End If
        End If
        
   
    Next cmpComponent


End Sub


Public Sub ImportVBAModules(ByRef wkb As Workbook, ByVal sFolder As String, _
    Optional ByVal PrefixForImports As String = "")
    
'------------------------------------------------------------------------------------------------------------------------
'   Code inspired by Ron De Bruin and Chip Pearson:
'   https://www.rondebruin.nl/win/s9/win002.htm
'   http://www.cpearson.com/excel/vbe.aspx
'
'
'   Requires references
'    - Microsoft Visual Basic For Applications Extensibility 5.3
'    - Microsoft Scripting Runtime
'
'   Requires Trust Access to VBA module:
'   In Excel 2007-2013, click the Developer tab and then click the Macro Security item. In that dialog,
'   choose Macro Settings and check the Trust access to the VBA project object model
'
'   Be aware that above may trigger action from antivirus software
'
'------------------------------------------------------------------------------------------------------------------------
    
'Imports VBA code sFolder
'If PrefixForImports is set only modules with that prefix are imported

    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim sTargetWorkbook As String
    Dim sImportPath As String
    Dim zFileName As String
    Dim PrefixOkforImport As Boolean
    Dim PrefixLength As Integer
    Dim cmpComponents As VBIDE.VBComponents

    If wkb.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If
    
    If wkb.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(sFolder).Files.Count = 0 Then
       Exit Sub
    End If


    Set cmpComponents = wkb.VBProject.VBComponents
    
    For Each objFile In objFSO.GetFolder(sFolder).Files
    
        PrefixLength = Len(PrefixForImports)
        PrefixOkforImport = PrefixForImports = "" Or _
            Left(objFile.Name, PrefixLength) = PrefixForImports
    
        If PrefixOkforImport And _
            ( _
                (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
                (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
                (objFSO.GetExtensionName(objFile.Name) = "bas") _
            ) Then
                cmpComponents.Import objFile.Path
        End If
        
    Next objFile
    
End Sub



Function DeleteModule(ByVal wkb As Workbook, ByVal sModuleName As String) As Boolean
    
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    
    On Error Resume Next
    Set VBProj = wkb.VBProject
    Set VBComp = VBProj.VBComponents(sModuleName)
    VBProj.VBComponents.Remove VBComp
    DeleteModule = (Err.Number = 0)
    On Error GoTo 0
    
End Function



Function RangeOfStoredData(ByVal ItemDescription As String) As Range

    Dim lo As ListObject
    Dim EvaluationFormula As String
    Dim ItemIndex As Integer
    
    Set lo = ThisWorkbook.Sheets("XL_Developer").ListObjects("tbl_Data")
    On Error Resume Next
    ItemIndex = WorksheetFunction.Match(ItemDescription, lo.ListColumns("Item").DataBodyRange, 0)
    
    If Err.Number <> 0 Then
        Set RangeOfStoredData = Nothing
    Else
        Set RangeOfStoredData = lo.ListColumns("Value").DataBodyRange.Cells(ItemIndex)
    End If

End Function


Function StoredDataValue(ByVal ItemDescription As String)

    Dim lo As ListObject
    Dim EvaluationFormula As String
    Dim ItemIndex As Integer
    
    Set lo = ThisWorkbook.Sheets("XL_Developer").ListObjects("tbl_Data")
    On Error Resume Next
    ItemIndex = WorksheetFunction.Match(ItemDescription, lo.ListColumns("Item").DataBodyRange, 0)
    
    If Err.Number <> 0 Then
        StoredDataValue = "NULL"
    Else
        StoredDataValue = lo.ListColumns("Value").DataBodyRange.Cells(ItemIndex).Value
    End If

End Function

