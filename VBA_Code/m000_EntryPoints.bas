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

Public Type TypeReportFieldSettings
    CubeFieldName As String
    FieldType As String
    Orientation As String
    Format As String
    CustomFormat As String
    Subtotal As Boolean
    SubtotalAtTop As Boolean
    BlankLine As Boolean
    FilterType As String
    FilterValues() As String
    CollapseFieldValues() As String
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




Sub CreateBiSpreadsheet()

    Dim wkb As Workbook
    Dim i As Integer
    
    'Setup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    'Create workbook and ensure it consists of only one sheet
    Set wkb = Application.Workbooks.Add
    If wkb.Sheets.Count <> 1 Then
        For i = wkb.Sheets.Count To 2 Step -1
            wkb.Sheets(i).Delete
        Next i
    End If
    
    CreateParameterSheet wkb
    CreateValidationSheet wkb
    CreateReportListSheet wkb
    CreateDataAccessQueriesPerReport wkb
    CreateDataLoadQueriesPerReport wkb
    CreateReportPropertiesSheet wkb
    CreateReportFieldSettingsSheet wkb
    CreateModelMeasuresSheet wkb
    CreateModelColumnsSheet wkb
    CreateModelCalculatedColumnsSheet wkb
    CreateModelRelationshipsSheet wkb
    CreateMissingLookupsSheet wkb
    CopyPowerQueriesBetweenFiles ThisWorkbook, wkb

    'Create index page and cleanup
    InsertIndexPage wkb
    wkb.Activate
    wkb.Sheets("Index").Activate
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

End Sub




'Sub AddValidationToReportFields()
'
'    Dim asMeasureList() As String
'    Dim asColumnList() As String
'    Dim sValidationString As String
'    Dim i As Integer
'    Dim lo As ListObject
'
'
'
'    sValidationString = ""
'    GetModelMeasureNames asMeasureList
'    GetModelColumnNames asColumnList
'
'    If ArrayIsDimensioned(asMeasureList) Then
'        For i = LBound(asMeasureList) To UBound(asMeasureList)
'            If sValidationString = "" Then
'                sValidationString = asMeasureList(i)
'            Else
'                sValidationString = sValidationString & "," & asMeasureList(i)
'            End If
'        Next i
'    End If
'
'    If ArrayIsDimensioned(asColumnList) Then
'        For i = LBound(asColumnList) To UBound(asColumnList)
'            If sValidationString = "" Then
'                sValidationString = asColumnList(i)
'            Else
'                sValidationString = sValidationString & "," & asColumnList(i)
'            End If
'        Next i
'    End If
'
'    If sValidationString <> "" Then
'        Set lo = ActiveWorkbook.Sheets("ReportFieldSettings").ListObjects("tbl_ReportFields")
'
'        On Error Resume Next
'        lo.ListColumns("Cube Field Name").DataBodyRange.Validation.Delete
'        On Error GoTo 0
'
'        lo.ListColumns("Cube Field Name").DataBodyRange.Validation.Add _
'            Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=sValidationString
'    End If
'
'
'    Application.ScreenUpdating = True
'    Application.EnableEvents = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.DisplayAlerts = True
'
'End Sub


Sub GenerateReports()

    Dim bValidSettings As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim ReportList() As TypeReportList
    Dim ReportProperties As TypeReportProperties
    Dim ReportFieldSettings() As TypeReportFieldSettings
    Dim pvt As PivotTable
    Dim sht As Worksheet

    'Setup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    GetReportList ReportList
    
    For i = 0 To UBound(ReportList)
        With ReportList(i)
            If .RunWithoutRefresh <> "" Or .RunWithRefresh <> "" Then
                'Get report setup
                GetReportProperties .ReportName, ReportProperties
                GetReportFieldSettings .ReportName, ReportFieldSettings
                
                'Create Report
                CreatePivotTable .SheetName, pvt
                CustomisePivotTable pvt, ReportProperties
                SetPivotFields pvt, ReportFieldSettings
                
                'Format and populate valuues on report sheet
                Set sht = ActiveWorkbook.Sheets(.SheetName)
                sht.Rows("1:5").Insert Shift:=xlDown
                sht.Name = .SheetName
                FormatSheet sht
                sht.Range("SheetHeading") = .ReportName
                sht.Range("SheetCategory") = .ReportCategory
                
            End If
        End With
    Next i

    InsertIndexPage ActiveWorkbook
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

End Sub

Sub WritesMeasuresColumnsRelationshipsToSheetsEntryPoint()

    
    'Setup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    WritesMeasuresColumnsRelationshipsToSheets
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    

End Sub



Sub GenerateMissingLookupTableItems()

    Dim i As Integer
    Dim j As Integer
    Dim loReports As ListObject
    Dim loDataLoadQueries As ListObject
    Dim loTableRelationships As ListObject
    Dim sReportName As String
    Dim sDataLoadQueryName As String
    Dim colLoadedQueryNames As Collection
    Dim item As Variant
    Dim sDaxStr As String
    Dim sSheetName As String
    Dim iCol As Integer

    'Setup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    WritesMeasuresColumnsRelationshipsToSheets
    Set loReports = ActiveWorkbook.Sheets("ReportList").ListObjects("tbl_ReportList")
    Set loDataLoadQueries = ActiveWorkbook.Sheets("DataLoadQueriesPerReport").ListObjects("tbl_DataLoadQueriesPerReport")
    Set loTableRelationships = ActiveWorkbook.Sheets("ModelRelationships").ListObjects("tbl_ModelRelationships")
    Set colLoadedQueryNames = New Collection
    
    
    'Saves the unique list of data queries (for the given reports that are selected for running)
    With loReports
        For i = 1 To .DataBodyRange.Rows.Count
            If .ListColumns("Run with table refresh").DataBodyRange.Cells(i) <> "" Or .ListColumns("Run without table refresh").DataBodyRange.Cells(i) <> "" Then
                sReportName = .ListColumns("Report Name").DataBodyRange.Cells(i)
                For j = 1 To loDataLoadQueries.DataBodyRange.Rows.Count
                    If loDataLoadQueries.ListColumns("Report Name").DataBodyRange.Cells(j) = sReportName Then
                        sDataLoadQueryName = loDataLoadQueries.ListColumns("Data Load Query Name").DataBodyRange.Cells(j)
                        On Error Resume Next
                        colLoadedQueryNames.Add item:=sDataLoadQueryName, Key:=sDataLoadQueryName
                        On Error GoTo 0
                    End If
                Next j
            End If
        Next i
    End With

    'Create Dax query sheet with missing items in lookup
    iCol = 2
    On Error Resume Next
    ActiveWorkbook.Sheets("MissingLookups").Delete
    On Error GoTo 0
    CreateMissingLookupsSheet ActiveWorkbook
    With loTableRelationships
        For Each item In colLoadedQueryNames
            For i = 1 To .DataBodyRange.Rows.Count
                If .ListColumns("Foreign Key Table").DataBodyRange.Cells(i) = item Then
                    
                    'Construct DAX query string
                    sDaxStr = "EVALUATE " & vbCrLf & "EXCEPT(" & vbCrLf & "    VALUES("
                    sDaxStr = sDaxStr & item
                    sDaxStr = sDaxStr & "[" & .ListColumns("Foreign Key Column").DataBodyRange.Cells(i) & "]), " & vbCrLf
                    sDaxStr = sDaxStr & "    VALUES(" & .ListColumns("Primary Key Table").DataBodyRange.Cells(i)
                    sDaxStr = sDaxStr & "[" & .ListColumns("Primary Key Column").DataBodyRange.Cells(i) & "]) " & vbCrLf
                    sDaxStr = sDaxStr & "    )"
                    
                    ActiveWorkbook.Sheets("MissingLookups").Cells(5, iCol) = sDaxStr
                    CreateDaxQueryTable sDaxStr, ActiveWorkbook.Sheets("MissingLookups").Cells(7, iCol)
                    iCol = iCol + 2
                            
                End If
            Next i
        Next item
    End With
    
    'Retrospectively set column widths - it seems akward to undo the width setting on the query table
    For i = iCol - 2 To 2 Step -2
        ActiveWorkbook.Sheets("MissingLookups").Columns(i).ColumnWidth = 60
    Next i
    ActiveWorkbook.Sheets("MissingLookups").Move After:=ActiveWorkbook.Sheets("ModelRelationships")
    InsertIndexPage ActiveWorkbook

    'Exit
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True


End Sub



