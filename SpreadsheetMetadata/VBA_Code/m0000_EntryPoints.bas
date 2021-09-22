Attribute VB_Name = "m0000_EntryPoints"
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

Sub InsertIndexPageActiveWorkbook()
    
    Dim IndexSheet As Worksheet

    StandardEntry
    Set IndexSheet = InsertIndexPage(ActiveWorkbook)
    IndexSheet.Activate
    IndexSheet.Range("DefaultCursorLocation").Select
    Set IndexSheet = Nothing
    StandardExit

End Sub


Function InsertReportingSheetSheetIntoActiveWorkbook()
    
    Dim ReportSht As ReportingSheet
    Dim wkb As Workbook
    Dim ReportSheetFormat As Dictionary

    StandardEntry

    Set wkb = ActiveWorkbook
    Set ReportSht = New ReportingSheet
    ReportSht.Create wkb, ActiveSheet.Index
    
    Set ReportSheetFormat = GetSavedReportSheetFormat
    ReportSht.SheetFont = ReportSheetFormat.item("Sheet font")
    ReportSht.DefaultFontSize = ReportSheetFormat.item("Default font size")
    ReportSht.ZoomPercentage = ReportSheetFormat.item("Zoom percentage")
    ReportSht.HeadingFontColour = Array( _
        ReportSheetFormat.item("Heading colour red (0 to 255)"), _
        ReportSheetFormat.item("Heading colour green (0 to 255)"), _
        ReportSheetFormat.item("Heading colour blue (0 to 255)"))
    ReportSht.HeadingFontSize = ReportSheetFormat.item("Heading font size")
    
    InsertIndexPage ActiveWorkbook
    ReportSht.Sheet.Activate
    ReportSht.DefaultCursorLocation.Select

    Set wkb = Nothing
    Set ReportSht = Nothing
    StandardExit

End Function


Sub ConvertActiveSheetToReportingSheet()

    Dim ReportSht As ReportingSheet
    Dim ReportSheetFormat As Dictionary

    StandardEntry
    Set ReportSht = New ReportingSheet

    ReportSht.CreateFromExistingSheet ActiveSheet
    Set ReportSheetFormat = GetSavedReportSheetFormat
    ReportSht.SheetFont = ReportSheetFormat.item("Sheet font")
    ReportSht.DefaultFontSize = ReportSheetFormat.item("Default font size")
    ReportSht.ZoomPercentage = ReportSheetFormat.item("Zoom percentage")
    ReportSht.HeadingFontColour = Array( _
        ReportSheetFormat.item("Heading colour red (0 to 255)"), _
        ReportSheetFormat.item("Heading colour green (0 to 255)"), _
        ReportSheetFormat.item("Heading colour blue (0 to 255)"))
    ReportSht.HeadingFontSize = ReportSheetFormat.item("Heading font size")
    
    InsertIndexPage ActiveWorkbook
    ReportSht.Sheet.Activate
    StandardExit

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


Sub ExportPowerQueriesInActiveWorkbookToConsolidatedFile()

    ExportPowerQueriesToConsolidatedFile ActiveWorkbook
    
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


Sub RunTableLooperOnActiveSheet()

    Dim ReportSheetSource As ReportingSheet
    Dim ReportSheetConsol As ReportingSheet
    Dim bReportSheetAssigned As Boolean
    
    StandardEntry
    Set ReportSheetSource = New ReportingSheet
    bReportSheetAssigned = ReportSheetSource.AssignExistingSheet(ActiveSheet)
    
    If Not bReportSheetAssigned Then
        MsgBox ("Not a valid sheet for table looping")
        GoTo Exitpoint
    End If
    
    If Not IsTableLooperSheet(ReportSheetSource.Sheet) Then
        MsgBox ("Not a valid sheet for table looping")
        GoTo Exitpoint
    End If
    
    Set ReportSheetConsol = InsertConsolLooperSheet(ReportSheetSource)
    LoopSourceAndCopyToConsolSheet ReportSheetSource, ReportSheetConsol
    FilterOutExcludedItems ReportSheetConsol
    SetLoopTableAndSheetFormat ReportSheetSource, ReportSheetConsol

Exitpoint:
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
       GoTo Exitpoint
    End If

    sFilePathAndName = fDialog.SelectedItems(1)
    sWorkbookName = fso.GetFileName(sFilePathAndName)


    If sWorkbookName = ActiveWorkbook.Name Then
        MsgBox ("Cannot copy between 2 workbooks with the same name, exiting...")
        GoTo Exitpoint
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

Exitpoint:
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




Sub ToggleErrorCheckRangeVisbilityOnSelectedSheets()
Attribute ToggleErrorCheckRangeVisbilityOnSelectedSheets.VB_ProcData.VB_Invoke_Func = "H\n14"

    Dim sht As Worksheet
    Dim ReportSheet As ReportingSheet
    Dim ReportIsAssigned As Boolean
    Dim obj As Object
    Dim ShowHiddenRange As Boolean
    Dim CurrentlyActiveSheet As Worksheet
    Dim IsFirstReportingSheetInSelection As Boolean
    
    StandardEntry
    Set CurrentlyActiveSheet = ActiveSheet
    IsFirstReportingSheetInSelection = True
    
    'Toggling can occur for multiple selected sheets
    'Visibility is set based on status of the first sheet
    For Each obj In ActiveWindow.SelectedSheets
        Set sht = obj
        Set ReportSheet = New ReportingSheet
        ReportIsAssigned = ReportSheet.AssignExistingSheet(sht)
        If ReportIsAssigned Then
            If IsFirstReportingSheetInSelection Then
                ShowHiddenRange = Not ReportSheet.HiddenRangesAreVisible
                IsFirstReportingSheetInSelection = False
            End If
            ReportSheet.ToggleErrorCheckRangeVisbility ShowHiddenRange
        End If
        Set ReportSheet = Nothing
    Next obj
    Set obj = Nothing
    CurrentlyActiveSheet.Activate
    StandardExit
    
End Sub



Sub IndexPageNavigation()
Attribute IndexPageNavigation.VB_ProcData.VB_Invoke_Func = "I\n14"

    Dim wkb As Workbook
    Dim TargetSheetName As String
    Dim TargetSheet As Worksheet

    Set wkb = ActiveWorkbook

    Select Case True
    
        Case ActiveSheet.Name <> "Index" And SheetExists(wkb, "Index")
            Sheets("Index").Activate
            On Error Resume Next
            On Error GoTo 0
            
        Case Selection.Rows.Count <> 1
            'Do Nothing
        
        Case wkb.Sheets("Index").Range("HiddenSheetNamesCol").Cells(Selection.Row) = ""
            'Do Nothing
            
        Case Else
            On Error Resume Next
            TargetSheetName = wkb.Sheets("Index").Range("HiddenSheetNamesCol").Cells(Selection.Row)
            Set TargetSheet = wkb.Sheets(TargetSheetName)
            TargetSheet.Activate
            On Error GoTo 0
            
    End Select

End Sub



Sub CreateRefencedPowerQueriesInActiveWorkbook()

    Dim FilePaths() As String
    Dim i As Integer
    Dim fso As FileSystemObject
    Dim wkb As Workbook
    Dim QueryText As String
    Dim QueryName As String
    
    Set fso = New FileSystemObject
    Set wkb = ActiveWorkbook
    GetPowerQueryFileNamesFromUser FilePaths
    For i = LBound(FilePaths) To UBound(FilePaths)
        QueryName = fso.GetBaseName(FilePaths(i))
        QueryText = PowerQueryReferencedToTextFile(FilePaths(i))
        wkb.Queries.Add QueryName, QueryText
    Next i

    MsgBox ("Queries created")

End Sub


Sub SetReportSheetFormat()

    Dim ReportSheetFormatStorage As zLIB_ListStorage
    Dim wkbUserInput As Workbook
    Dim UserInputSheet As Worksheet
    Dim UserInputListObj As ListObject

    StandardEntry
    Set ReportSheetFormatStorage = New zLIB_ListStorage
    ReportSheetFormatStorage.AssignStorage ThisWorkbook, "ReportSheetFormat"

    Set wkbUserInput = Application.Workbooks.Add
    Set UserInputSheet = wkbUserInput.Sheets(1)
    FormatSheet UserInputSheet
    UserInputSheet.Range("SheetHeading") = "Report Sheet Formating --> Run 'Format->" & _
        "Save default report sheet format"
    UserInputSheet.Range("SheetCategory") = ""
    
    ReportSheetFormatStorage.ListObj.Range.Copy
    UserInputSheet.Activate
    UserInputSheet.Range("B5").PasteSpecial xlPasteValues
    
    Set UserInputListObj = UserInputSheet.ListObjects.Add(xlSrcRange, Range("$B$5").CurrentRegion, , xlYes)
    FormatTable UserInputListObj
    UserInputListObj.Name = "tbl_ReportSheetFormat"

    ActiveWindow.WindowState = xlMaximized

Exitpoint:
    Set ReportSheetFormatStorage = Nothing
    StandardExit

End Sub


Sub SaveReportSheetFormat()

    Dim i As Integer
    Dim ReportSheetFormatStorage As zLIB_ListStorage
    Dim UserInputListObj As ListObject

    StandardEntry
    Set ReportSheetFormatStorage = New zLIB_ListStorage
    ReportSheetFormatStorage.AssignStorage ThisWorkbook, "ReportSheetFormat"
    Set UserInputListObj = ActiveWorkbook.Worksheets(1).ListObjects("tbl_ReportSheetFormat")
    
    For i = 1 To ReportSheetFormatStorage.NumberOfRecords
        ReportSheetFormatStorage.ListObj.ListColumns("Value").DataBodyRange.Cells(i) = _
            UserInputListObj.ListColumns("Value").DataBodyRange.Cells(i)
    Next i

    MsgBox ("Report sheet format updated")
    
    UserInputListObj.Parent.Parent.Close
    
    ThisWorkbook.Save
    StandardExit

End Sub




Function UpdateReportSheetFormatsInActiveWorkbook()
    
'    Dim wkb As Workbook
'    Dim SheetActiveAtSubStart As Worksheet
'    Dim sht As Worksheet
'    Dim ReportSheet As ReportingSheet
'    Dim ReportSheetAssigned As Boolean
'    Dim RptSheetFormat As TypeReportSheetFormat
'
'    StandardEntry
'    Set wkb = ActiveWorkbook
'    Set SheetActiveAtSubStart = ActiveSheet
'    ReadSavedReportSheetFormat RptSheetFormat
'
'    For Each sht In wkb.Worksheets
'        Set ReportSheet = New ReportingSheet
'        ReportSheetAssigned = ReportSheet.AssignExistingSheet(sht)
'        If ReportSheetAssigned And (sht.Visible = xlSheetVisible) Then
'            ApplyReportSheetFormat ReportSheet, RptSheetFormat
'        End If
'    Next sht
'
'    SheetActiveAtSubStart.Activate
'    Set wkb = Nothing
'    Set SheetActiveAtSubStart = Nothing
'    StandardExit

End Function



