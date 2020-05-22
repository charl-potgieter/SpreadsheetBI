Attribute VB_Name = "m000_EntryPoints"
Option Explicit

Public Type TypeReportList
    ReportName As String
    SheetName As String
    ReportCategory As String
    RunWithRefresh As String
    RunWithoutRefresh As String
End Type

Public Type TypeReportProperties
    AutoFit As Boolean
    RowTotals As Boolean
    ColumnTotals As Boolean
    DisplayExpandButtons As Boolean
    DisplayFieldHeaders As Boolean
End Type



Public Type TypeModelMeasures
    Name As String
    UniqueName As String
    Visible As Boolean
    Expression As String
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


Public Const MaxInt As Integer = 32767


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



Sub InsertFormattedSheetIntoActiveWorkbook()
    
    Dim sht As Worksheet
    
    Set sht = ActiveWorkbook.Sheets.Add(After:=ActiveSheet)
    FormatSheet sht

End Sub


Sub FormatActiveSheet()

    FormatSheet ActiveSheet

End Sub



Sub FormatHeadings()

    Application.ScreenUpdating = False
    
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
    
    
    Application.ScreenUpdating = True


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
    Dim FSO As FileSystemObject
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
            Set FSO = New FileSystemObject
            sPowerQueryName = Replace(FSO.GetFileName(sPowerQueryFilePath), ".m", "")
            ImportOrRefreshSinglePowerQuery sPowerQueryFilePath, sPowerQueryName, ActiveWorkbook
        Next i
    End If
    
    MsgBox ("Queries imported")

End Sub




Sub InsertIndexPageActiveWorkbook()


    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    InsertIndexPage ActiveWorkbook
        
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

End Sub






Sub FormatPivotTableFlatten()

    Dim pvt As PivotTable
    Dim pvtField As PivotField
    Dim b_mu As Boolean
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
        
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


    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True




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
    
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    
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

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True



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
    
    'Setup
    sActiveSheetName = ActiveSheet.Name
    sActiveRangeAddress = Selection.Address
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

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
    
    'Cleanup
    Worksheets(sActiveSheetName).Activate
    Range(sActiveRangeAddress).Select
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True


End Sub






Sub WriteModelInfoToSheets()
'Writes below information from power pivot model in activeworkbook to worksheets:
'   Model Measures
'   Model columns
'   Measure Relationships
    
    'Setup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    WriteModelMeasuresToSheet
    WriteModelCalcColsToSheet
    WriteModelColsToSheet
    WriteRelationshipsToSheet
    
    'Cleanup
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    

End Sub


Sub MimimiseRibbon()
Attribute MimimiseRibbon.VB_ProcData.VB_Invoke_Func = "R\n14"
    CommandBars.ExecuteMso "MinimizeRibbon"
End Sub

Sub CreatePowerQueryGeneratorSheet()
'Creates a sheet in active workbook to be utilsed for the generation of "hard coded" power query tables

    Dim iMsgBoxResponse As Integer

    'Setup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False


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

    'Cleanup
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True


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
    
    'Setup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    
    sQueryName = ActiveWorkbook.Sheets("PqTableGenerator").Range("TableName")

    If QueryExists(sQueryName, ActiveWorkbook) Then
        MsgBox ("Query with the same name already exists.  New query not generated")
        Exit Sub
    End If
    
    Set lo = ActiveWorkbook.Sheets("PqTableGenerator").ListObjects("tbl_PqTableGenerator")
    
        
    sQueryText = "let" & vbCr & vbCr & _
        "    tbl = Table.FromRecords({" & vbCr
        
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
                sQueryText = sQueryText & "], " & vbCr
            Else
                sQueryText = sQueryText & "]" & vbCrLf
            End If
        Next j
        If i = lo.DataBodyRange.Rows.Count Then
            sQueryText = sQueryText & "        }), " & vbCr & vbCr
        End If
    Next i
        
        
    'Changed Type portion of power query
    sQueryText = sQueryText & "    ChangedType = Table.TransformColumnTypes(" & vbCr & _
        "       tbl, " & vbCr & "        {" & vbCr

    For j = 1 To lo.HeaderRowRange.Cells.Count
        sQueryText = sQueryText & "            {""" & lo.HeaderRowRange.Cells(j) & """, " & lo.HeaderRowRange.Cells(j).Offset(-2, 0) & "}"
        If j <> lo.HeaderRowRange.Cells.Count Then
            sQueryText = sQueryText & "," & vbCr
        Else
            sQueryText = sQueryText & vbCr
        End If
    Next j
    sQueryText = sQueryText & vbCr & "        })" & vbCr & vbCr & _
        "in" & vbCr & "    ChangedType"

    ActiveWorkbook.Queries.Add sQueryName, sQueryText
    
    MsgBox ("Query Generated")

    'Cleanup
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True


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

    'Setup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    sMetaDataRootPath = ActiveWorkbook.Path & Application.PathSeparator & "SpreadsheetMetadata"
    sWorksheetStructurePath = sMetaDataRootPath & Application.PathSeparator & "WorksheetStructure"
    sPowerQueriesPath = sMetaDataRootPath & Application.PathSeparator & "PowerQueries"
    sVbaCodePath = sMetaDataRootPath & Application.PathSeparator & "VBA_Code"
    
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
    
    'Generate Worksheet structure metadata text files
    GenerateMetadataFileWorksheets ActiveWorkbook, sWorksheetStructurePath & Application.PathSeparator & "MetadataWorksheets.txt"
    GenerateMetadataFileListObjectFields ActiveWorkbook, sWorksheetStructurePath & Application.PathSeparator & "ListObjectFields.txt"
    GenerateMetadataFileListObjectValues ActiveWorkbook, sWorksheetStructurePath & Application.PathSeparator & "ListObjectFieldValues.txt"
    GenerateMetadataFileListObjectFormat ActiveWorkbook, sWorksheetStructurePath & Application.PathSeparator & "ListObjectFormat.txt"

    'Export VBA code
    ExportVBAModules ActiveWorkbook, sVbaCodePath

    'Export Power Queries
    ExportPowerQueriesToFiles sPowerQueriesPath, ActiveWorkbook

    'Cleanup
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
        
    MsgBox ("Metadata created")

End Sub


Sub CopyQueriesFromSpreadsheetBIIntoActiveWorkbook()

    CopyQueriesFromSpreadsheetBI ActiveWorkbook
    
End Sub
