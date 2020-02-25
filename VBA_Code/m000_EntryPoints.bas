Attribute VB_Name = "m000_EntryPoints"
Option Explicit

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
        For i = 2 To wkb.Sheets.Count
            wkb.Sheets(i).Delete
        Next i
    End If
    
    CreateParamaterSheet wkb
    CreateReportListSheet wkb
    CreateQueriesPerReportSheet wkb
    CreateReportPropertiesSheet wkb
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




Sub AddValidationToReportFields()
    
    Dim asMeasureList() As String
    Dim asColumnList() As String
    Dim sValidationString As String
    Dim i As Integer
    Dim lo As ListObject



    sValidationString = ""
    GetModelMeasureNames asMeasureList
    GetModelColumnNames asColumnList
        
    If ArrayIsDimensioned(asMeasureList) Then
        For i = LBound(asMeasureList) To UBound(asMeasureList)
            If sValidationString = "" Then
                sValidationString = asMeasureList(i)
            Else
                sValidationString = sValidationString & "," & asMeasureList(i)
            End If
        Next i
    End If
    
    If ArrayIsDimensioned(asColumnList) Then
        For i = LBound(asColumnList) To UBound(asColumnList)
            If sValidationString = "" Then
                sValidationString = asColumnList(i)
            Else
                sValidationString = sValidationString & "," & asColumnList(i)
            End If
        Next i
    End If
    
    If sValidationString <> "" Then
        Set lo = ActiveWorkbook.Sheets("ReportFields").ListObjects("tbl_ReportFields")
        
        On Error Resume Next
        lo.ListColumns("Cube Field Name").DataBodyRange.Validation.Delete
        On Error GoTo 0
        
        lo.ListColumns("Cube Field Name").DataBodyRange.Validation.Add _
            Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=sValidationString
    End If


    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

End Sub


Sub GenerateReports()

    Dim bValidSettings As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim loReportList As ListObject
    Dim loReportFields As ListObject
    Dim pvt As PivotTable
    

    bValidSettings = ReportSettingsAreValid
    If Not bValidSettings Then Exit Sub
    
    Set loReportList = ActiveWorkbook.Worksheets("ReportList").ListObjects("tbl_ReportList")
    Set loReportFields = ActiveWorkbook.Worksheets("ReportFields").ListObjects("tbl_ReportFields")
    
    With loReportList
        For i = 1 To .DataBodyRange.Rows.Count
            If .ListColumns("Run without table refresh").DataBodyRange.Cells(i) <> "" Then
                CreatePivotTable .ListColumns("Sheet Name").DataBodyRange.Cells(i), .ListColumns("Report Name").DataBodyRange.Cells(i), .ListColumns("Report Category"), pvt
                CustomisePivotTable .ListColumns("Report Name").DataBodyRange.Cells(i)
                With loReportFields
                    For j = 1 To .DataBodyRange.Rows.Count
                        If .ListColumns("Report Name").DataBodyRange.Cells(j) = loReportList.ListColumns("Report Name").DataBodyRange(i) Then
                            SetPivotFields pvt, .ListColumns("Cube Field Name").DataBodyRange.Cells(j), .ListColumns("Orientation").DataBodyRange.Cells(j)
                        End If
                    Next j
                End With
            End If
        Next i
    End With

    InsertIndexPageActiveWorkbook

End Sub
