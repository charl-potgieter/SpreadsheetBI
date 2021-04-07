Attribute VB_Name = "EntryPoints"
'@Folder ("000_EntryPoints")

Option Explicit

Public Type TypeModelMeasures
    Name As String
    UniqueName As String
    Visible As Boolean
    Expression As String
    Table As String
End Type

Public Type TypeModelColumns
    Name As String
    UniqueName As String
    TableName As String
    Visible As Boolean
End Type

Public Type TypeModelCalcColumns
    Name As String
    TableName As String
    Expression As String
End Type

Public Type TypeModelRelationship
    ForeignKeyTable As String
    ForeignKeyColumn As String
    PrimaryKeyTable As String
    PrimaryKeyColumn As String
    Active As Boolean
End Type

'Utilised for saving data
Public Type TypePowerReportStorageRecord
    SheetName As String
    Name As String
    DataType As String
    Property As String
    Value As String
    CubeFieldPosition As Variant
End Type


Public Enum EnumPivotReportType
    PowerPivotSource
    ExcelTableSource
    ExcelTableOnly
End Enum

Public Type TypePivotReportUserSelection
    SelectionMade As Boolean
    ReportList() As String
    PivotReportType As EnumPivotReportType
    GenerateValueCopyNewWorkbook As Boolean
    RetainLiveInCurrentWorkbook As Boolean
End Type

Public Const MaxInt As Integer = 32767
Public Const cPR_MaxStorageRecords As Long = 1000000  'PR  = PowerReport


Sub DisplayPopUpMenu()
Attribute DisplayPopUpMenu.VB_ProcData.VB_Invoke_Func = "M\n14"

    DeletePopUpMenu
    CreatePopUpMenu
    Application.CommandBars(gcsMenuName).ShowPopup

End Sub


Sub FormatZeroDecimalNumberFormat()
    SetNumberFormat "#,##0_);(#,##0);-??"
End Sub



Sub FormatOneDecimalNumberFormat()
    SetNumberFormat "#,##0.0_);(#,##0.0);-??"
End Sub



Sub FormatTwoDecimalsNumberFormat()
    SetNumberFormat "#,##0.00_);(#,##0.00);-??"
End Sub



Sub FormatTwoDigitPercentge()
    SetNumberFormat "0.00%"
End Sub


Sub FormatFourDigitPercentge()
    SetNumberFormat "0.0000%"
End Sub


Sub FormatDate()
    SetNumberFormat "dd-mmm-yy"
End Sub


Sub FormatDashboardIconStyle()
'Creates custom formatting to displat different dashboard style icons, for positive, negative, zero and text values
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Note that Hex character codes are obtained by using excel menu, insert -> symbol
'(select font as arial top right, subset as geometric shape, select hex code bottom left.
'Note also use of _) and * below to get good alignment of symbols (the bracket seems to give enough space to align triangles with diamond)
'Sheet Font should be set to Calibri to ensure best display (which is as per the m010_General.FormatSheet sub in this workbook)
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color 10] " & ChrW(&H25B2) & "_);" & _
        "[Red] " & ChrW(&H25BC) & "_);" & _
        "[Color 46] " & ChrW(&H2666) & " ;" & _
        "[Blue] * " & ChrW(&H25BA) & "_ "
    
End Sub


Sub FormatZeroDecimalAndArrows()
'Custom formatting displays numbers and up and down arrows as appropriate
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Note that Hex character codes are obtained by using excel menu, insert -> symbol
'(select font as arial top right, subset as geometric shape, select hex code bottom left.
'Note also use of _) and * below to get good alignment of symbols (the bracket seems to give enough space to align triangles with diamond)
'Sheet Font should be set to Calibri to ensure best display (which is as per the m010_General.FormatSheet sub in this workbook)
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color10]#,##0_) " & ChrW(&H25B2) & "_);" & _
        "[Red] (#,##0) " & ChrW(&H25BC) & "_);" & _
        "-????;" & _
        "General"
    
End Sub

Sub FormatOneDecimalAndArrow()
'Custom formatting displays numbers and up and down arrows as appropriate
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Note that Hex character codes are obtained by using excel menu, insert -> symbol
'(select font as arial top right, subset as geometric shape, select hex code bottom left.
'Note also use of _) and * below to get good alignment of symbols (the bracket seems to give enough space to align triangles with diamond)
'Sheet Font should be set to Calibri to ensure best display (which is as per the m010_General.FormatSheet sub in this workbook)
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color10]#,##0.0_) " & ChrW(&H25B2) & "_);" & _
        "[Red] (#,##0.0) " & ChrW(&H25BC) & "_);" & _
        "-????;" & _
        "General"
    
End Sub

Sub FormatTwoDecimalsAndArrow()
'Custom formatting displays numbers and up and down arrows as appropriate
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Note that Hex character codes are obtained by using excel menu, insert -> symbol
'(select font as arial top right, subset as geometric shape, select hex code bottom left.
'Note also use of _) and * below to get good alignment of symbols (the bracket seems to give enough space to align triangles with diamond)
'Sheet Font should be set to Calibri to ensure best display (which is as per the m010_General.FormatSheet sub in this workbook)
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color10]#,##0.00_) " & ChrW(&H25B2) & "_);" & _
        "[Red] (#,##0.00) " & ChrW(&H25BC) & "_);" & _
        "-????;" & _
        "General"
    
End Sub

Sub FormatZeroDigitPercentageAndArrow()
'Custom formatting displays numbers and up and down arrows as appropriate
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Note that Hex character codes are obtained by using excel menu, insert -> symbol
'(select font as arial top right, subset as geometric shape, select hex code bottom left.
'Note also use of _) and * below to get good alignment of symbols (the bracket seems to give enough space to align triangles with diamond)
'Sheet Font should be set to Calibri to ensure best display (which is as per the m010_General.FormatSheet sub in this workbook)
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color10] 0% " & ChrW(&H25B2) & "_);" & _
        "[Red] -0% " & ChrW(&H25BC) & "_);" & _
        "0%??;" & _
        "General"
    
End Sub

Sub FormatTwoDigitPercentageAndArrow()
'Custom formatting displays numbers and up and down arrows as appropriate
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Note that Hex character codes are obtained by using excel menu, insert -> symbol
'(select font as arial top right, subset as geometric shape, select hex code bottom left.
'Note also use of _) and * below to get good alignment of symbols (the bracket seems to give enough space to align triangles with diamond)
'Sheet Font should be set to Calibri to ensure best display (which is as per the m010_General.FormatSheet sub in this workbook)
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color10] 0.00% " & ChrW(&H25B2) & "_);" & _
        "[Red] -0.00% " & ChrW(&H25BC) & "_);" & _
        "0.00%??;" & _
        "General"
    
End Sub

Sub FormatFourDigitPercentageAndArrow()
'Custom formatting displays numbers and up and down arrows as appropriate
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Note that Hex character codes are obtained by using excel menu, insert -> symbol
'(select font as arial top right, subset as geometric shape, select hex code bottom left.
'Note also use of _) and * below to get good alignment of symbols (the bracket seems to give enough space to align triangles with diamond)
'Sheet Font should be set to Calibri to ensure best display (which is as per the m010_General.FormatSheet sub in this workbook)
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color10] 0.0000% " & ChrW(&H25B2) & "_);" & _
        "[Red] -0.0000% " & ChrW(&H25BC) & "_);" & _
        "0.0000%??;" & _
        "General"
    
End Sub

Sub FormatOkError()
'1 Displays OK in green, zero ERROR in red.  Negatives adn text are hidden
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color10]OK ;;[Red]\E\R\RO\R;"

End Sub



Function InsertFormattedSheetIntoActiveWorkbook()
    Dim ReportSht As ReportingSheet
    Set ReportSht = New ReportingSheet
    ReportSht.Create ActiveWorkbook, ActiveSheet.Index
End Function


Sub FormatActiveSheet()

    Dim Report As ReportingSheet
    Set Report = New ReportingSheet
    Report.CreateFromExistingSheet ActiveSheet

End Sub



Sub FormatHeadings()

    StandardEntry
    
    'Remove all current borders
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    'Set new borders
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    'Set header colour
    With Selection.Interior
        .Color = RGB(217, 225, 242)
        .Pattern = xlSolid
    End With
    
    Selection.Font.Bold = True
    
    'Set Text allignment
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    
    StandardExit

End Sub



Sub ExportPowerQueriesInActiveWorkbookToFiles()

    Dim sFolderSelected As String
    
    sFolderSelected = GetFolder
    If NumberOfFilesInFolder(sFolderSelected) <> 0 Then
        MsgBox ("Please select an empty folder...exiting")
        Exit Sub
    End If
    
    ExportPowerQueriesToFiles sFolderSelected, ActiveWorkbook
    MsgBox ("Queries Exported")

End Sub


Sub ExportNonStandardPowerQueriesInActiveWorkbookToFiles()
'Exports power queries without fn_std or template_std prefix

    Dim sFolderSelected As String
    
    sFolderSelected = GetFolder
    If NumberOfFilesInFolder(sFolderSelected) <> 0 Then
        MsgBox ("Please select an empty folder...exiting")
        Exit Sub
    End If
    
    ExportNonStandardPowerQueriesToFiles sFolderSelected, ActiveWorkbook
    MsgBox ("Queries Exported")

End Sub


Sub ImportPowerQueriesFromSelectedFolderNonRecursive()

    Dim sFolderSelected As String
    
    sFolderSelected = GetFolder
    ImportOrRefreshPowerQueriesInFolder sFolderSelected, False
    MsgBox ("Queries imported")
    
End Sub



Sub ImportPowerQueriesFromSelectedFolderRecursive()

    Dim sFolderSelected As String
    
    sFolderSelected = GetFolder
    ImportOrRefreshPowerQueriesInFolder sFolderSelected, True
    MsgBox ("Queries imported")
    
End Sub



Sub ImportSelectedPowerQueries()
'Requires reference to Microsoft Scripting runtime library

    Dim sPowerQueryFilePath As String
    Dim sPowerQueryName As String
    Dim fDialog As FileDialog
    Dim fso As FileSystemObject
    Dim i As Integer
    
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With fDialog
        .AllowMultiSelect = True
        .Title = "Select power query / queries"
        .InitialFileName = ThisWorkbook.Path
        .Filters.Clear
        .Filters.Add "m Power Query Files", "*.m"
    End With
    
    'fDialog.Show value of -1 below means success
    If fDialog.Show = -1 Then
        For i = 1 To fDialog.SelectedItems.Count
            sPowerQueryFilePath = fDialog.SelectedItems(i)
            Set fso = New FileSystemObject
            sPowerQueryName = Replace(fso.GetFileName(sPowerQueryFilePath), ".m", "")
            ImportOrRefreshSinglePowerQuery sPowerQueryFilePath, sPowerQueryName, ActiveWorkbook
        Next i
    End If
    
    MsgBox ("Queries imported")

End Sub


Sub InsertIndexPageActiveWorkbook()
    StandardEntry
    InsertIndexPage ActiveWorkbook
    StandardExit
End Sub



Sub FormatPivotTableFlatten()

    Dim pvt As PivotTable
    Dim pvtField As PivotField
    Dim b_mu As Boolean
    
    StandardEntry
    On Error Resume Next
    Set pvt = ActiveCell.PivotTable
    On Error GoTo 0
    
    If Not pvt Is Nothing Then
        With pvt
            'Get update status and suspend updates
            b_mu = .ManualUpdate
            .ManualUpdate = True
            
            .RowAxisLayout xlTabularRow
            .ColumnGrand = True
            .RowGrand = True
            .HasAutoFormat = False
            .ShowDrillIndicators = False
            
            For Each pvtField In pvt.PivotFields
                If pvtField.Orientation = xlRowField Then
                    pvtField.RepeatLabels = True
                    pvtField.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                End If
            Next pvtField
            
            'Restore update status
            .ManualUpdate = b_mu
            
        End With
    End If

    StandardExit
End Sub


Sub FormatActiveTable()

    FormatTable ActiveCell.ListObject
    
End Sub






Sub CreateDataModelRelationships()
'Creates relationships based on listobject in sheet Model_Relationhsips which has the below fields:
'ID, Foreign Key Table, Foreign Key Column, Primary Key Table, Primary Key Column, Active


    Dim lo As ListObject
    Dim i As Integer
    Dim mdl As Model
    Dim sForeignKeyTable As String
    Dim sForeignKeyCol As String
    Dim sPrimaryKeyTable As String
    Dim sPrimaryKeyCol As String
    
    StandardEntry
    Set mdl = ActiveWorkbook.Model
    Set lo = ActiveWorkbook.Sheets("Model_Relationships").ListObjects(1)
    For i = 1 To lo.DataBodyRange.Rows.Count
        sForeignKeyTable = lo.ListColumns("Foreign Key Table").DataBodyRange.Cells(i)
        sForeignKeyCol = lo.ListColumns("Foreign Key Column").DataBodyRange.Cells(i)
        sPrimaryKeyTable = lo.ListColumns("Primary Key Table").DataBodyRange.Cells(i)
        sPrimaryKeyCol = lo.ListColumns("Primary Key Table").DataBodyRange.Cells(i)
        mdl.ModelRelationships.Add _
            mdl.ModelTables(sForeignKeyTable).ModelTableColumns(sForeignKeyCol), _
            mdl.ModelTables(sPrimaryKeyTable).ModelTableColumns(sPrimaryKeyCol)
    Next i
    StandardExit
End Sub



Sub TableLooper()

'Implements a looping mechanism:
' - that loops over various keys
' - populates a calculation table utilising each of the above keys
' - copies to a consolidated output sheet
'
'Precondition
'-------------
'Listobject tbl_LoopController exists on sheet LoopController
'The listobject has the below fields
' - item
' - value
' - notes
'The below values appear in the item field
' - Index range
' - Input key
' - Input calculation table
' - Target sheet name
' - Target sheet inserted after
' - target sheet heading
' - target sheet category


    Dim Arr
    Dim i As Integer
    Dim sht As Worksheet
    Dim dblRowToPaste As Double
    Dim loOutput As ListObject
    Dim sActiveSheetName As String
    Dim sActiveRangeAddress As String
    Dim rngTableInputKey As Range
    Dim sTargetSheetName As String
    Dim sAfterSheet As String
    Dim sSheetHeading As String
    Dim sSheetCategory As String
    Dim loCalc As ListObject
    
    Const iStartTableRow As Integer = 5
    Const iStartTableCol As Integer = 2
    
    
    StandardEntry
    sActiveSheetName = ActiveSheet.Name
    sActiveRangeAddress = Selection.Address

    'Read inputs for looping function
    Arr = WorksheetFunction.Transpose(Range(LooperValue("Index Range")))
    Set rngTableInputKey = Range(LooperValue("Input Key"))
    Set loCalc = Range(LooperValue("Input Calculation Table")).ListObject
    sTargetSheetName = LooperValue("Target Sheet Name")
    sAfterSheet = LooperValue("Target Sheet Inserted After")
    sSheetHeading = LooperValue("Target Sheet Heading")
    sSheetCategory = LooperValue("Target Sheet Category")
    
    
    'Create sheet for consolidated output of calculations
    If SheetExists(ActiveWorkbook, sTargetSheetName) Then
        ActiveWorkbook.Sheets(sTargetSheetName).Delete
    End If
    Set sht = ActiveWorkbook.Sheets.Add(After:=Worksheets(sAfterSheet))
    FormatSheet sht
    sht.Name = sTargetSheetName
    sht.Range("SheetHeading") = sSheetHeading
    sht.Range("A1") = sSheetCategory
    
    loCalc.HeaderRowRange.Copy
    sht.Cells(iStartTableRow, iStartTableCol).PasteSpecial xlPasteValues

    For i = LBound(Arr) To UBound(Arr)
        rngTableInputKey = Arr(i)
        Application.CalculateFull
        Application.Wait Now + #12:00:01 AM#
        loCalc.DataBodyRange.Copy
        dblRowToPaste = iStartTableRow + sht.Cells(iStartTableRow, iStartTableCol).CurrentRegion.Rows.Count
        sht.Cells(dblRowToPaste, iStartTableCol).PasteSpecial xlPasteValues
    Next i

    Set loOutput = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Cells(iStartTableRow, iStartTableCol).CurrentRegion, XlListObjectHasHeaders:=xlYes)
    loOutput.Name = "tbl_" & sTargetSheetName
    
    FormatTable loOutput
    sht.Select
    Rows("6:6").Select
    ActiveWindow.FreezePanes = True
    
    InsertIndexPageActiveWorkbook
    
    Worksheets(sActiveSheetName).Activate
    Range(sActiveRangeAddress).Select
    StandardExit

End Sub



Sub WriteModelInfoToSheets()
'Writes below information from power pivot model in activeworkbook to worksheets:
'   Model Measures
'   Model columns
'   Measure Relationships
    
    Dim iMsgBoxResponse As Integer
    
    StandardEntry
        
    If SheetExists(ActiveWorkbook, "ModelMeasures") Or SheetExists(ActiveWorkbook, "ModelCalcColumns") Or _
    SheetExists(ActiveWorkbook, "ModelColumns") Or SheetExists(ActiveWorkbook, "ModelRelationships") Then
            iMsgBoxResponse = MsgBox("Sheets already exists, delete?", vbQuestion + vbYesNo + vbDefaultButton2)
            If iMsgBoxResponse = vbNo Then
                Exit Sub
            End If
    End If
    
    WriteModelMeasuresToSheet
    WriteModelCalcColsToSheet
    WriteModelColsToSheet
    WriteModelRelationshipsToSheet
    
    StandardExit
End Sub


Sub MimimiseRibbon()
    CommandBars.ExecuteMso "MinimizeRibbon"
End Sub

Sub CreatePowerQueryGeneratorSheet()
'Creates a sheet in active workbook to be utilsed for the generation of "hard coded" power query tables
'utilising sub GeneratePowerQueryTable

    Dim iMsgBoxResponse As Integer

    StandardEntry
    'Give user choice to delete sheet or cancel if sheet already exists?
    If SheetExists(ActiveWorkbook, "PqTableGenerator") Then
        iMsgBoxResponse = MsgBox("Sheet already exists, delete?", vbQuestion + vbYesNo + vbDefaultButton2)
        If iMsgBoxResponse = vbNo Then
            Exit Sub
        Else
            ActiveWorkbook.Sheets("PqTableGenerator").Delete
        End If
    End If
        
    CreateTableGeneratorSheet ActiveWorkbook
    StandardExit
End Sub


Sub GeneratePowerQueryTable()
'Generates a hard coded power query table from ActiveWorkbook.Sheets("TableGenerator").ListObjects("tbl_TableGenerator")
'Query name is as per defined name on sheet entitled "TableName"
'Column types are stored in cells 2 rows above the tbl_Generator table

    Dim sQueryName As String
    Dim sQueryText As String
    Dim i As Integer
    Dim j As Integer
    Dim lo As ListObject
    
    StandardEntry
    sQueryName = ActiveWorkbook.Sheets("PqTableGenerator").Range("TableName")

    If QueryExists(sQueryName, ActiveWorkbook) Then
        MsgBox ("Query with the same name already exists.  New query not generated")
        Exit Sub
    End If
    
    Set lo = ActiveWorkbook.Sheets("PqTableGenerator").ListObjects("tbl_PqTableGenerator")
    
        
    sQueryText = "let" & vbCrLf & vbCrLf & _
        "    tbl = Table.FromRecords({" & vbCrLf
        
    'Table from records portion of power query
    For i = 1 To lo.DataBodyRange.Rows.Count
        For j = 1 To lo.HeaderRowRange.Cells.Count
            If j = 1 Then
                sQueryText = sQueryText & "        ["
            End If
            sQueryText = sQueryText & _
                lo.HeaderRowRange.Cells(j) & " = """ & lo.ListColumns(j).DataBodyRange.Cells(i) & """"
            If j <> lo.HeaderRowRange.Cells.Count Then
                sQueryText = sQueryText & ", "
            ElseIf i <> lo.DataBodyRange.Rows.Count Then
                sQueryText = sQueryText & "], " & vbCrLf
            Else
                sQueryText = sQueryText & "]" & vbCrLf
            End If
        Next j
        If i = lo.DataBodyRange.Rows.Count Then
            sQueryText = sQueryText & "        }), " & vbCrLf & vbCrLf
        End If
    Next i
        
        
    'Changed Type portion of power query
    sQueryText = sQueryText & "    ChangedType = Table.TransformColumnTypes(" & vbCrLf & _
        "       tbl, " & vbCrLf & "        {" & vbCrLf

    For j = 1 To lo.HeaderRowRange.Cells.Count
        sQueryText = sQueryText & "            {""" & lo.HeaderRowRange.Cells(j) & """, " & lo.HeaderRowRange.Cells(j).Offset(-2, 0) & "}"
        If j <> lo.HeaderRowRange.Cells.Count Then
            sQueryText = sQueryText & "," & vbCrLf
        Else
            sQueryText = sQueryText & vbCrLf
        End If
    Next j
    sQueryText = sQueryText & vbCrLf & "        })" & vbCrLf & vbCrLf & _
        "in" & vbCrLf & "    ChangedType"

    ActiveWorkbook.Queries.Add sQueryName, sQueryText
    
    MsgBox ("Query Generated")
    StandardExit

End Sub



Sub ExportVBAcodeExModuleName()
'Exports VBA code into path of current workbook.  The module name is excluded so that the code
'can simply be copied and pasted into the VBA ide, rather than imported.
'Likely that this functionality will only be used for spreadsheets with only a single module of code

    ExportVBAModules ActiveWorkbook, ActiveWorkbook.Path, True
    MsgBox ("VBA Code exported")

End Sub



Sub GenerateSpreadsheetMetaData()

'Generates selected spreadsheet data to allo the spreadsheet to be recreated
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
    Dim sWorksheetStructurePath As String
    Dim sPowerQueriesPath As String
    Dim sVbaCodePath As String
    Dim sDataModelPath As String

    StandardEntry

    sMetaDataRootPath = ActiveWorkbook.Path & Application.PathSeparator & "SpreadsheetMetadata"
    sWorksheetStructurePath = sMetaDataRootPath & Application.PathSeparator & "WorksheetStructure"
    sPowerQueriesPath = sMetaDataRootPath & Application.PathSeparator & "PowerQueries"
    sVbaCodePath = sMetaDataRootPath & Application.PathSeparator & "VBA_Code"
    sDataModelPath = sMetaDataRootPath & Application.PathSeparator & "DataModel"
    
    'Rather ask user to manually delete rather than have risky folder deletions in VBA code
    If FolderExists(sMetaDataRootPath) Then
        MsgBox ("Manually delete " & sMetaDataRootPath & " before continuing.  Exiting")
        Exit Sub
    End If
    
    'Create folders for storing metadata
    CreateFolder sMetaDataRootPath
    CreateFolder sWorksheetStructurePath
    CreateFolder sPowerQueriesPath
    CreateFolder sVbaCodePath
    CreateFolder sDataModelPath
    
    'Generate Worksheet structure metadata text files
    GenerateMetadataFileWorksheets ActiveWorkbook, sWorksheetStructurePath & Application.PathSeparator & "MetadataWorksheets.txt"
    GenerateMetadataFileListObjectFields ActiveWorkbook, sWorksheetStructurePath & Application.PathSeparator & "ListObjectFields.txt"
    GenerateMetadataFileListObjectValues ActiveWorkbook, sWorksheetStructurePath & Application.PathSeparator & "ListObjectFieldValues.txt"
    GenerateMetadataFileListObjectFormat ActiveWorkbook, sWorksheetStructurePath & Application.PathSeparator & "ListObjectFormat.txt"

    'Generate power pivot data model text files
    WriteModelMeasuresToPipeDelimtedText ActiveWorkbook, sDataModelPath & Application.PathSeparator & "Measures.txt"
    WriteModelCalcColsToPipeDelimitedFile ActiveWorkbook, sDataModelPath & Application.PathSeparator & "CalculatedColumns.txt"
    WriteModelColsToPipeDelimitedFile ActiveWorkbook, sDataModelPath & Application.PathSeparator & "Columns.txt"
    WriteModelRelationshipsToPipeDelimitedFile ActiveWorkbook, sDataModelPath & Application.PathSeparator & "Relationships.txt"
    WriteModelMeasuresToHumanReadableText ActiveWorkbook, sDataModelPath & Application.PathSeparator & "MeasuresHumanReadable.dax"

    'Export VBA code
    ExportVBAModules ActiveWorkbook, sVbaCodePath

    'Export Power Queries
    ExportPowerQueriesToFiles sPowerQueriesPath, ActiveWorkbook

    MsgBox ("Metadata created")
    StandardExit
    

End Sub


Sub CopyPowerQueriesFromWorkbook()
'Copies power queries from selected workbook into active workbook

    
    Dim fDialog As FileDialog, Result As Integer
    Dim sFilePathAndName As String
    Dim bWorkbookIsOpen As Boolean
    Dim fso As New FileSystemObject
    Dim sWorkbookName As String
    Dim wkbSource As Workbook
    Dim wkbTarget As Workbook
    Dim qry As WorkbookQuery

    StandardEntry
    
    Set wkbTarget = ActiveWorkbook
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
        
    'Get file from file picker
    fDialog.AllowMultiSelect = False
    fDialog.InitialFileName = ActiveWorkbook.Path
    fDialog.Filters.Clear
    fDialog.Filters.Add "Excel files", "*.xlsx, *.xlsm"
    fDialog.Filters.Add "All files", "*.*"
     
    'Exit sub  if no file is selected
    If fDialog.Show <> -1 Then
       GoTo ExitPoint
    End If
     
    sFilePathAndName = fDialog.SelectedItems(1)
    sWorkbookName = fso.GetFileName(sFilePathAndName)
    
    
    If sWorkbookName = ActiveWorkbook.Name Then
        MsgBox ("Cannot copy between 2 workbooks with the same name, exiting...")
        GoTo ExitPoint
    End If
    
    
    'Open source workbook if not open
    If WorkbookIsOpen(sWorkbookName) Then
        bWorkbookIsOpen = True
        Set wkbSource = Workbooks(sWorkbookName)
    Else
        bWorkbookIsOpen = False
        Set wkbSource = Application.Workbooks.Open(sFilePathAndName)
    End If
    
    'Copy queries from source to active workbook
    For Each qry In wkbSource.Queries
        If QueryExists(qry.Name, wkbTarget) Then
            wkbTarget.Queries(qry.Name).Formula = qry.Formula
        Else
            wkbTarget.Queries.Add qry.Name, qry.Formula
        End If
    Next qry
    
    'Close source workbook if it was not open before this sub was run
    If Not bWorkbookIsOpen Then
        wkbSource.Close
    End If
    
    wkbTarget.Activate
    MsgBox ("Power Queries copied")

ExitPoint:
    StandardExit

End Sub






Sub TempDeleteAllPQ()
'Deletes all power queries in active workbook

    Dim qry As WorkbookQuery
    
    StandardEntry
    For Each qry In ThisWorkbook.Queries
        qry.Delete
    Next qry
    StandardExit

End Sub






Sub AssignPivotReportFormulaStorageInActiveWorkbook()
'Utilsied as storage for saving of DAX query files utilised to create a drillable pivot report
'that uses excel data table (potentially with formulas) fo backing

    Dim bStorageCreated As Boolean
    StandardEntry
    DataPivotReporting.AssignPivotReportFormulaStorage ActiveWorkbook
    StandardExit

End Sub


Sub SavePivotReportMetadataInActiveWorkbook()
'Reads all pivot table metadata in active workbook and saves on worksheets in active workbook

    Dim sht As Worksheet
    Dim bValidAssignment As Boolean
    Dim PvtReport As PivotReport
    Dim vStorageObject As Variant 'abstract away the storage structure

    StandardEntry
    Set vStorageObject = AssignPivotReportStructureStorage(ActiveWorkbook, True)

    For Each sht In ActiveWorkbook.Worksheets
        Set PvtReport = New PivotReport
        bValidAssignment = PvtReport.AssignToExistingSheet(sht)
        If bValidAssignment Then
            SaveSinglePivotReportDataToStorage vStorageObject, PvtReport
        End If
        Set PvtReport = Nothing
    Next sht
    
    StandardExit

End Sub


Sub CreatePivotReportFromMetadata()

    Dim vStorageObject As Variant
    Dim UserReportSelection As TypePivotReportUserSelection
    Dim i As Long
    Dim PvtReport As PivotReport
    Dim sReportName As String

    StandardEntry
    Set vStorageObject = DataPivotReporting.AssignPivotReportStructureStorage(ActiveWorkbook)
    
    'Exit if no report metadata exists on active sheet
    If vStorageObject Is Nothing Then
        MsgBox ("No pivot report metadata exists on active sheet")
        GoTo ExitPoint
    End If
    
    UserReportSelection = GetUserPivotReportSelection
    
    With UserReportSelection
        If .SelectionMade = False Then GoTo ExitPoint
        Select Case .PivotReportType
            Case PowerPivotSource
                For i = LBound(.ReportList) To UBound(.ReportList)
                    sReportName = .ReportList(i)
                    Set PvtReport = New PivotReport
                    PvtReport.CreateEmptyPivotReport ActiveWorkbook, sReportName, DataModel
                    DesignPivotReportBasedOnStoredData vStorageObject, PvtReport
                Next i
            Case ExcelTableOnly
                'TODO Create Excel Table
            Case ExcelTableSource
                'TODO Create Excel Table and pivot
        End Select
    End With
    
ExitPoint:
    StandardExit

End Sub


'
'Sub CreatePowerPivotReportFromMetaData()
'
'    Dim Report As PivotReport
'    Dim vStorageObject As Variant
'    Dim Records_PivotTableProperties() As TypePowerReportStorageRecord
'    Dim Records_CubeFieldOrientationProperties() As TypePowerReportStorageRecord
'    Dim Records_CubeFieldExOrientationProperties() As TypePowerReportStorageRecord
'    Dim Records_PivotFieldSubtotalProperties() As TypePowerReportStorageRecord
'    Dim Records_PivotFieldProperties() As TypePowerReportStorageRecord
'    Dim iPivotFirstRowNumber As Integer
'    Dim sSheetHeading As String
'    Dim sSheetCategory As String
'    Dim vFreezeDetails As Variant
'    Dim vRowRangeColWidths As Variant
'    Dim vDataBodyRangeColWidth As Variant
'
'    Dim v As Variant
'    Dim item As Variant
'    Dim uf As ufPivotReportGenerator
'    Dim i As Integer
'    Dim wkbReportSource As Workbook
'    Dim wkbValues As Workbook
'    Dim sSheetName As String
'
'
'
'    'Setup
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'    Application.Calculation = xlCalculationManual
'    Application.DisplayAlerts = False
'
'    Set wkbReportSource = ActiveWorkbook
'    Set vStorageObject = DataPivotReporting.AssignPivotReportStorage(ActiveWorkbook)
'
'    'Exit if no report metadata exists on active sheet
'    If vStorageObject Is Nothing Then
'        MsgBox ("No report metadata exists on active sheet")
'        Exit Sub
'    End If
'
'    Set uf = New ufPivotReportGenerator
'
'    uf.Show
'
'    If Not uf.bCancelled Then
'        If uf.chkValueCopy.Value = True Then
'            Set wkbValues = Workbooks.Add
'        End If
'        For i = 0 To uf.lbReports.ListCount - 1
'            If uf.lbReports.Selected(i) Then
'
'                'Get required data to generate the report
'                sSheetName = uf.lbReports.List(i)
'                DataPivotReporting.PR_GetPivotTableProperties vStorageObject, sSheetName, _
'                    Records_PivotTableProperties
'                DataPivotReporting.PR_GetPivotCubeFieldDataOrientationSortedByCubeFieldPosition _
'                    vStorageObject, sSheetName, Records_CubeFieldOrientationProperties
'                DataPivotReporting.PR_GetPivotCubeFieldDataPropertiesExOrientation _
'                    vStorageObject, sSheetName, Records_CubeFieldExOrientationProperties
'                DataPivotReporting.PR_GetPivotFieldDataSubtotalProperty vStorageObject, _
'                    sSheetName, Records_PivotFieldSubtotalProperties
'                DataPivotReporting.PR_GetPivotFieldDataPropertiesExSubtotals vStorageObject, _
'                    sSheetName, Records_PivotFieldProperties
'                iPivotFirstRowNumber = DataPivotReporting.PR_GetFirstPivotRow(vStorageObject, _
'                    sSheetName)
'                sSheetHeading = DataPivotReporting.PR_GetHeadingBasedOnSheetName( _
'                    vStorageObject, sSheetName)
'                sSheetCategory = DataPivotReporting.PR_GetCategoryBasedOnSheetName( _
'                    vStorageObject, sSheetName)
'                vFreezeDetails = DataPivotReporting.PR_GetFreezePaneLocation(vStorageObject, _
'                    sSheetName)
'                vRowRangeColWidths = DataPivotReporting.PR_GetRowRangeColWidths(vStorageObject, _
'                    sSheetName)
'                vDataBodyRangeColWidth = DataPivotReporting.PR_GetDataBodyRowRangeColWidth _
'                    (vStorageObject, sSheetName)
'
'                'Create Report
'                Set Report = New PivotReport
'
'                Select Case True
'
'                    Case uf.obPowePivotSource.Value = True
'                        Report.CreatePowerPivotReportFromData wkbReportSource, sSheetName, _
'                            Records_PivotTableProperties, Records_CubeFieldOrientationProperties, _
'                            Records_CubeFieldExOrientationProperties, Records_PivotFieldSubtotalProperties, _
'                            Records_PivotFieldProperties, iPivotFirstRowNumber, sSheetHeading, _
'                            sSheetCategory, vFreezeDetails, vRowRangeColWidths, vDataBodyRangeColWidth
'                        'Copy Values if selected
'                        If uf.chkValueCopy.Value = True Then
'                            Report.CreateValueCopy wkbValues
'                            If uf.chkRetainLiveReport = False Then
'                                Report.Delete
'                            End If
'                        End If
'
'                    Case uf.obExcelTableOnly.Value = True
'                        Report.CreateExcelBacking wkbReportSource, sSheetName, sSheetCategory
'
'
'                End Select
'
'            End If
'        Next i
'    End If
'
'    Unload uf
'    Set uf = Nothing
'
'    wkbValues.Saved = True
'
''ExitPoint:
'    Application.ScreenUpdating = True
'    Application.EnableEvents = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.DisplayAlerts = True
'
'
'End Sub


