Attribute VB_Name = "m000_EntryPoints"
Option Explicit

Sub AAAACreateDisplayPopUpMenu()
Attribute AAAACreateDisplayPopUpMenu.VB_ProcData.VB_Invoke_Func = "M\n14"
    
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

End Sub




Sub ImportPowerQueriesFromSelectedFolderNonRecursive()

    Dim sFolderSelected As String
    
    sFolderSelected = GetFolder
    ImportOrRefreshPowerQueriesInFolder sFolderSelected, False

End Sub



Sub ImportPowerQueriesFromSelectedFolderRecursive()

    Dim sFolderSelected As String
    
    sFolderSelected = GetFolder
    ImportOrRefreshPowerQueriesInFolder sFolderSelected, True

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


    Dim arr
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
    arr = WorksheetFunction.Transpose(Range(LooperValue("Index Range")))
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

    For i = LBound(arr) To UBound(arr)
        rngTableInputKey = arr(i)
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
    Dim sht As Worksheet
    Dim lo As ListObject
    
    'Setup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    
    Set wkb = Application.Workbooks.Add
    
    'Ensure workbook consists of only one sheet
    If wkb.Sheets.Count <> 1 Then
        For i = 1 To wkb.Sheets.Count
            wkb.Sheets(i).Delete
        Next i
    End If
    
    'Create parameter sheet
    Set sht = wkb.Sheets(1)
    FormatSheet sht
    sht.Name = "Parameters"
    sht.Range("SheetHeading") = "Parameters"
    sht.Range("SheetCategory") = "Setup"
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:C8"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_Parameters"
        .HeaderRowRange.Cells(1) = "Parameter"
        .HeaderRowRange.Cells(2) = "Value"
        .ListColumns("Parameter").DataBodyRange.Cells(1) = "Date_Start"
        .ListColumns("Parameter").DataBodyRange.Cells(2) = "Date_End"
        .ListColumns("Value").DataBodyRange.Cells.NumberFormat = "dd-mmm-yy"
        .HeaderRowRange.RowHeight = .HeaderRowRange.RowHeight * 2
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 30
    sht.Range("C:C").ColumnWidth = 60

    'Create report list sheet
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ReportList"
    sht.Range("SheetHeading") = "Report List"
    sht.Range("SheetCategory") = "Setup"
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:D7"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_ReportList"
        .HeaderRowRange.Cells(1) = "Report Name"
        .HeaderRowRange.Cells(2) = "Run with table refresh"
        .HeaderRowRange.Cells(3) = "Run without table refresh"
        .HeaderRowRange.RowHeight = .HeaderRowRange.RowHeight * 2
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 60
    sht.Range("C:C").ColumnWidth = 30
    sht.Range("D:D").ColumnWidth = 30
    sht.Range("B4") = "Clear data from non-dependent tables (mark with X)"
    With sht.Range("D4")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(217, 225, 242)
    End With
    SetOuterBorders sht.Range("D4")

    
    'Create queries per report sheet
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "QueriesPerReport"
    sht.Range("SheetHeading") = "Queries per report"
    sht.Range("SheetCategory") = "Setup"
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:D8"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_QueriesPerReport"
        .HeaderRowRange.Cells(1) = "Report Name"
        .HeaderRowRange.Cells(2) = "Report selected for run and query refresh"
        .HeaderRowRange.Cells(3) = "Query Name"
        .HeaderRowRange.RowHeight = .HeaderRowRange.RowHeight * 2
        .ListColumns("Report selected for run and query refresh").DataBodyRange.Formula = _
             "=(COUNTIFS(tbl_ReportList[Report Name], [@[Report Name]], tbl_ReportList[Run with table refresh], "" * "")) > 0"
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 50
    sht.Range("C:C").ColumnWidth = 30
    sht.Range("D:D").ColumnWidth = 50
    
    
    'Create report properties sheet
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ReportProperties"
    sht.Range("SheetHeading") = "Report properties"
    sht.Range("SheetCategory") = "Setup"
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:D8"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_ReportProperties"
        .HeaderRowRange.Cells(1) = "Report Name"
        .HeaderRowRange.Cells(2) = "AutoFit"
        .HeaderRowRange.Cells(3) = "Total Rows"
        .HeaderRowRange.Cells(3) = "Total Columns"
        .HeaderRowRange.RowHeight = .HeaderRowRange.RowHeight * 2
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 50
    sht.Range("C:C").ColumnWidth = 30
    sht.Range("D:D").ColumnWidth = 30
    sht.Range("E:E").ColumnWidth = 30
    
    
    'Create report report fields sheet
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ReportFields"
    sht.Range("SheetHeading") = "Report fields"
    sht.Range("SheetCategory") = "Setup"
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:D8"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_ReportFields"
        .HeaderRowRange.Cells(1) = "Report Name"
        .HeaderRowRange.Cells(2) = "Cube Field Name"
        .HeaderRowRange.Cells(3) = "Orientation"
        .HeaderRowRange.Cells(3) = "Format"
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 50
    sht.Range("C:C").ColumnWidth = 50
    sht.Range("D:D").ColumnWidth = 30
    sht.Range("E:E").ColumnWidth = 30
    
        
    'Copy Power Queries from this Workbook into the newly created workbook
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
