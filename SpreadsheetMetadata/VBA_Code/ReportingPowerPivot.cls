VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReportingPowerPivot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Storage.Reporting"
Option Explicit

'-----------------------------------------------------------------------------
'   Requires reference to Microsoft Scripting runtime
'-----------------------------------------------------------------------------

'Use "This" declaration as an easy way to get intellisense to the classes private variables
'https://rubberduckvba.wordpress.com/2020/02/27/vba-classes-gateway-to-solid/
Private Type TypePowerReport
    Name As String
    pvt As PivotTable
    ReportSht As ReportingSheet
End Type
    
Private this As TypePowerReport


Public Function AssignToExistingSheet(ByRef sht As Worksheet) As Boolean

    Dim TestReportSht As ReportingSheet
    Dim bTestReportSheetAssigned As Boolean
    
    Set TestReportSht = New ReportingSheet
    bTestReportSheetAssigned = TestReportSht.AssignExistingSheet(sht)
    
    Select Case True
        Case Not bTestReportSheetAssigned
            AssignToExistingSheet = False 'Only allow pivot reports to exist on reportsheets
            Set TestReportSht = Nothing
        Case sht.PivotTables.Count <> 1
            AssignToExistingSheet = False
        Case sht.PivotTables(1).PivotCache.WorkbookConnection.Name <> "ThisWorkbookDataModel"
            AssignToExistingSheet = False
        Case HasFilters(sht.PivotTables(1))
            AssignToExistingSheet = False
        Case Else
            AssignToExistingSheet = True
            Set this.pvt = sht.PivotTables(1)
            Set this.ReportSht = New ReportingSheet
            this.ReportSht.AssignExistingSheet sht
            this.Name = this.ReportSht.Heading
    End Select
        
End Function

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(ByVal sName As String)
    this.Name = sName
    this.ReportSht.Heading = sName
End Property

Public Property Get SheetProperties() As Dictionary
    Dim ReturnDictionary As Dictionary
    Set ReturnDictionary = New Dictionary
    With ReturnDictionary
        .Add "FirstCellAddress", this.pvt.TableRange2.Cells(1).Address
        .Add "SheetName", this.pvt.Parent.Name
        .Add "Category", this.ReportSht.Category
        .Add "Heading", this.ReportSht.Heading
        .Add "FreezePaneRow", this.ReportSht.FreezePaneRow
        .Add "FreezePaneCol", this.ReportSht.FreezePaneCol
    End With
    Set SheetProperties = ReturnDictionary
End Property


Public Property Let SheetProperties(ByVal Properties As Dictionary)
'Pivot table is moved from original sheet to new sheet when sheet properties are set
    
    Dim shtOriginalSheet  As Worksheet
    Dim wkb As Workbook
    Dim sFirstCellAddress As String
    
    Set shtOriginalSheet = this.pvt.Parent
    Set wkb = shtOriginalSheet.Parent
    Set this.ReportSht = New ReportingSheet
    With this.ReportSht
        .Create wkb, shtOriginalSheet.Index
        .Name = Properties.item("SheetName")
        .Category = Properties.item("Category")
        .Heading = Properties.item("Heading")
        .FreezePaneRow = Properties.item("FreezePaneRow")
        .FreezePaneCol = Properties.item("FreezePaneCol")
        
        sFirstCellAddress = Properties.item("FirstCellAddress")
        this.pvt.TableRange2.Cut
        .Sheet.Paste .Sheet.Range(sFirstCellAddress)
        'if below step is ignored the this.pvt reference is deleted on cut and paste
        Set this.pvt = .Sheet.Range(sFirstCellAddress).PivotTable
        shtOriginalSheet.Delete
        
    End With
End Property


Public Property Get PivotTableProperties() As Dictionary

    Dim PropertyNames As Variant
    Dim PropertyName As Variant
    Dim PropertyValue As Variant
    Dim ReturnDictionary As Dictionary

    PropertyNames = Array("Name", "LayoutRowDefault", "PageFieldWrapCount", "CompactRowIndent", _
        "PageFieldOrder", "CompactLayoutColumnHeader", "GrandTotalName", "TableStyle2", "Value", _
        "CompactLayoutRowHeader", "AllowMultipleFilters", "DisplayEmptyColumn", "DisplayEmptyRow", _
        "DisplayErrorString", "EnableDataValueEditing", "FieldListSortAscending", _
        "InGridDropZones", "ManualUpdate", "MergeLabels", "PrintDrillIndicators", "PrintTitles", _
        "ShowTableStyleColumnStripes", "ShowTableStyleRowStripes", "ShowValuesRow", "SmallGrid", _
        "VisualTotalsForSets", "CalculatedMembersInFilters", "ColumnGrand", "DisplayContextTooltips", _
        "DisplayFieldCaptions", "DisplayImmediateItems", "DisplayMemberPropertyTooltips", "DisplayNullString", _
        "EnableDrilldown", "EnableFieldDialog", "EnableFieldList", "EnableWizard", "HasAutoFormat", "PreserveFormatting", _
        "RepeatItemsOnEachPrintedPage", "RowGrand", "ShowDrillIndicators", "ShowPageMultipleItemLabel", _
        "ShowTableStyleColumnHeaders", "ShowTableStyleLastColumn", "ShowTableStyleRowHeaders", "SortUsingCustomLists", _
        "SubtotalHiddenPageItems", "TotalsAnnotation", "ViewCalculatedMembers", "VisualTotals", "AlternativeText", _
        "ErrorString", "NullString", "PageFieldStyle", "Summary", "VacatedStyle")

        Set ReturnDictionary = New Dictionary
        For Each PropertyName In PropertyNames
            PropertyValue = CallByName(this.pvt, PropertyName, VbGet)
            ReturnDictionary.Add PropertyName, PropertyValue
        Next PropertyName

        Set PivotTableProperties = ReturnDictionary

End Property


Public Property Let PivotTableProperties(Properties As Dictionary)

    Dim Property As Variant
    For Each Property In Properties.Keys
        CallByName this.pvt, Property, VbLet, Properties.item(Property)
    Next Property

End Property


'Public Property Get CubeFields() As CubeFields
'    Set CubeFields = this.pvt.CubeFields
'End Property
'
'
'Public Property Get PivotFields() As PivotFields
'    Set PivotFields = this.pvt.PivotFields
'End Property
'
'Public Property Get CubeFieldsOrientation() As Dictionary
'
'    Dim ReturnDictionary As Dictionary
'    Dim PvtCubeField As CubeField
'
'    Set ReturnDictionary = New Dictionary
'    For Each PvtCubeField In this.pvt.CubeFields
'        If PvtCubeField.Orientation <> xlHidden Then
'            ReturnDictionary.Add PvtCubeField.Name, PvtCubeField.Orientation
'        End If
'    Next PvtCubeField
'    Set CubeFieldsOrientation = ReturnDictionary
'
'End Property
'
'
'
'Public Property Let CubeFieldsOrientation(OrientationDictionary As Dictionary)
'
'    Dim CubeFieldName As Variant
'
'    For Each CubeFieldName In OrientationDictionary
'        this.pvt.CubeFields(CubeFieldName).Orientation = OrientationDictionary.item(CubeFieldName)
'    Next CubeFieldName
'
'End Property



Public Property Get CubeFieldProperties() As Dictionary

    Dim ReturnDictionary As Dictionary
    Dim PvtCubeField As CubeField
    Dim PropertyNames As Variant
    Dim PropertyName As Variant
    Dim CubePositionNameProperty As Variant
    Dim PropertyValue As Variant
    Dim sPaddedPosition As String
    
    PropertyNames = Array("Orientation", "Caption")
    Set ReturnDictionary = New Dictionary
    For Each PvtCubeField In this.pvt.CubeFields
        If PvtCubeField.Orientation <> xlHidden Then
        
            'Position is handled seperately as cannot be set by its property
            '(Fields need to be added in order)
            sPaddedPosition = Right("00" & PvtCubeField.Position, 3)
            For Each PropertyName In PropertyNames
                CubePositionNameProperty = "Position=" & sPaddedPosition & "|" & _
                    PvtCubeField.Name & "|" & PropertyName
                PropertyValue = CallByName(PvtCubeField, PropertyName, VbGet)
                ReturnDictionary.Add CubePositionNameProperty, PropertyValue
            Next PropertyName
        End If
    Next PvtCubeField
    Set CubeFieldProperties = ReturnDictionary

End Property


Property Let CubeFieldProperties(PropertyDictionary As Dictionary)

    Dim CubePositionNameProperty As Variant
    Dim sCubeFieldName As String
    Dim sProperty
    Dim sPropertyValue
    
    'Need to set Orientation first as other properties follow
    'Property Dictionary is sorted in display field order to ensure correct
    'order in pivot table as position can't seem to be set with VBA
    For Each CubePositionNameProperty In PropertyDictionary
        sCubeFieldName = Split(CubePositionNameProperty, "|")(1)
        sProperty = Split(CubePositionNameProperty, "|")(2)
        sPropertyValue = PropertyDictionary.item(CubePositionNameProperty)
        If sProperty = "Orientation" Then
            CallByName this.pvt.CubeFields(sCubeFieldName), sProperty, _
                VbLet, sPropertyValue
        End If
    Next CubePositionNameProperty
    
    
    'Set non-Orientation properties (Orientation needs to be set first as above)
    For Each CubePositionNameProperty In PropertyDictionary
        sCubeFieldName = Split(CubePositionNameProperty, "|")(1)
        sProperty = Split(CubePositionNameProperty, "|")(2)
        sPropertyValue = PropertyDictionary.item(CubePositionNameProperty)
        If sProperty <> "Orientation" Then
            CallByName this.pvt.CubeFields(sCubeFieldName), sProperty, _
                VbLet, sPropertyValue
        End If
    Next CubePositionNameProperty

End Property



Public Property Get PivotFieldProperties() As Dictionary
'Some of Object model properties are set at pivot field level.   Others at a cubefield level

    Dim ReturnDictionary As Dictionary
    Set ReturnDictionary = New Dictionary
    Dim pvtField As PivotField

    For Each pvtField In this.pvt.PivotFields
        Select Case True
            Case pvtField.Name = "Values"
                'Special case.  This is only case where orientation is set at a pivot
                'field level.  Other cases are set at cube field level
                    ReturnDictionary.Add pvtField.Name & "|Orientation", pvtField.Orientation
            Case pvtField.Orientation = xlDataField
                'For some unknown reason pvtField.NumberFormat does yield correct results
                ReturnDictionary.Add pvtField.Name & "|NumberFormat", this.pvt.PivotFields(pvtField.Name).NumberFormat
            Case (pvtField.Orientation = xlRowField) Or (pvtField.Orientation = xlColumnField)
                ReturnDictionary.Add pvtField.Name & "|LayoutBlankLine", pvtField.LayoutBlankLine
                ReturnDictionary.Add pvtField.Name & "|LayoutCompactRow", pvtField.LayoutCompactRow
                ReturnDictionary.Add pvtField.Name & "|LayoutForm", pvtField.LayoutForm
                ReturnDictionary.Add pvtField.Name & "|LayoutPageBreak", pvtField.LayoutPageBreak
                ReturnDictionary.Add pvtField.Name & "|RepeatLabels", pvtField.RepeatLabels
                ReturnDictionary.Add pvtField.Name & "|Subtotals", pvtField.Subtotals(1)
                If pvtField.Subtotals(1) = True Then
                    ReturnDictionary.Add pvtField.Name & "|SubtotalName", pvtField.SubtotalName
                    ReturnDictionary.Add pvtField.Name & "|LayoutSubtotalLocation", pvtField.LayoutSubtotalLocation
                End If
        End Select
    Next pvtField

    Set PivotFieldProperties = ReturnDictionary

End Property


Property Let PivotFieldProperties(PropertyDictionary As Dictionary)

    Dim FieldNameAndProperty As Variant
    Dim sPivotFieldName As String
    Dim sProperty
    Dim sPropertyValue
    
    'Subtotals need to be set first due to flow on impact.  note also that this property
    'is indexed
    For Each FieldNameAndProperty In PropertyDictionary
        sPivotFieldName = Split(FieldNameAndProperty, "|")(0)
        sProperty = Split(FieldNameAndProperty, "|")(1)
        sPropertyValue = PropertyDictionary.item(FieldNameAndProperty)
        If sProperty = "Subtotals" Then
            this.pvt.PivotFields(sPivotFieldName).Subtotals(1) = sPropertyValue
        End If
    Next FieldNameAndProperty
    
    For Each FieldNameAndProperty In PropertyDictionary
        sPivotFieldName = Split(FieldNameAndProperty, "|")(0)
        sProperty = Split(FieldNameAndProperty, "|")(1)
        sPropertyValue = PropertyDictionary.item(FieldNameAndProperty)
        If sProperty <> "Subtotals" Then
            CallByName this.pvt.PivotFields(sPivotFieldName), sProperty, _
                VbLet, sPropertyValue
            
        End If
    Next FieldNameAndProperty

End Property



Property Get ColumnWidths() As Dictionary
    
    Dim ReturnDictionary As Dictionary
    Set ReturnDictionary = New Dictionary
    Dim sRowRangeColWidthsStr As String
    Dim rng As Range

    'All data body col widths are set the same width
    If PivotTableHasDataBodyRange Then
        ReturnDictionary.Add "DataBodyRangeColWidth", _
            this.pvt.DataBodyRange.Columns(1).ColumnWidth
    End If
    
    'Col widths of area to left of DataBodyRange ("RowRange") stored as pipe delimited string
    If PivotTableHasRowRange Then
        sRowRangeColWidthsStr = ""
        For Each rng In this.pvt.RowRange.Rows(1).Cells
            If sRowRangeColWidthsStr = "" Then
                sRowRangeColWidthsStr = CStr(rng.ColumnWidth)
            Else
                sRowRangeColWidthsStr = sRowRangeColWidthsStr & "|" & CStr(rng.ColumnWidth)
            End If
        Next rng
        ReturnDictionary.Add "RowRangeColWidths", sRowRangeColWidthsStr
    End If
    
    Set ColumnWidths = ReturnDictionary
    
End Property



Property Let ColumnWidths(PropertyDictionary As Dictionary)

    Dim aRowRangeColWidths() As String
    Dim i As Long
    
    If PropertyDictionary.Exists("DataBodyRangeColWidth") Then
        this.pvt.DataBodyRange.Columns.ColumnWidth = PropertyDictionary.item("DataBodyRangeColWidth")
    End If
    
    If PropertyDictionary.Exists("RowRangeColWidths") Then
        aRowRangeColWidths = Split(PropertyDictionary.item("RowRangeColWidths"), "|")
        For i = LBound(aRowRangeColWidths) To UBound(aRowRangeColWidths)
            this.pvt.TableRange2.Columns(i + 1).ColumnWidth = aRowRangeColWidths(i)
        Next i
    End If
    

End Property




'Public Property Get Sundry() As Dictionary
'    Dim ReturnDictionary As Dictionary
'    Set ReturnDictionary = New Dictionary
'    ReturnDictionary.Add "FirstCellAddress", this.pvt.TableRange2.Cells(1).Address
'    Set Sundry = ReturnDictionary
'End Property



'Property Let Sundry(InputDictionary As Dictionary)
'
'    Dim sht As Worksheet
'    Dim sFirstCellAddress As String
'
'    sFirstCellAddress = InputDictionary.item("FirstCellAddress")
'    Set sht = this.pvt.Parent
'    this.pvt.TableRange2.Cut
'    sht.Paste sht.Range(sFirstCellAddress)
'
'End Property




Public Sub CreateEmptyPowerPivotReport(ByVal wkb As Workbook, ByVal sReportName As String)

    Dim sht As Worksheet
    Const sInitialFirstCell As String = "A1"

    Set sht = wkb.Sheets.Add(After:=wkb.ActiveSheet)
    Set this.pvt = wkb.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
        wkb.Connections("ThisWorkbookDataModel"), Version:=6). _
        CreatePivotTable(sht.Range(sInitialFirstCell))
    this.Name = sReportName
    
End Sub


Private Function HasFilters(ByVal pvtTable As PivotTable) As Boolean
'Returns true if any of the Cubefields are filtered

    Dim PvtCubeField As CubeField

    HasFilters = False
    For Each PvtCubeField In pvtTable.CubeFields
        If Not PvtCubeField.AllItemsVisible Then
            HasFilters = True
        End If
    Next PvtCubeField

End Function


Private Function PivotTableHasRowRange() As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = this.pvt.RowRange
    PivotTableHasRowRange = (Err.Number = 0)
    On Error GoTo 0
End Function


Private Function PivotTableHasDataBodyRange() As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = this.pvt.DataBodyRange
    PivotTableHasDataBodyRange = (Err.Number = 0)
    On Error GoTo 0
End Function


'Private Function DaxQueryFileExists() As Boolean
''Requires reference to Microsoft Scripting runtime
''An alternative solution exists using the DIR function but this seems to result in memory leak and file is
''not released by VBA
'
'    Dim fso As Scripting.FileSystemObject
'    Dim sFolderPath As String
'
'    sFolderPath = ActiveWorkbook.Path & Application.PathSeparator & _
'        csDaxSubFolder & Application.PathSeparator & this.ReportSht.Name & ".dax"
'
'    Set fso = New Scripting.FileSystemObject
'
'    DaxQueryFileExists = fso.FileExists(sFolderPath)
'    Set fso = Nothing
'
'End Function



'
'Property Get SundryProperties() As Long
'    TODO get the below properties here:
'     - FirstPivotTableCellAddress
'     - DataBodyRangeColWidth
'     - RowRangeColWidths
'    FirstPivotTableRow = this.pvt.TableRange2.Cells(1).Row
'End Property




'Public Sub CreatePowerPivotReportFromData( _
'    ByRef wkbReportSource As Workbook, _
'    ByVal sSheetName As String, _
'    ByRef Records_PivotTableProperties() As TypePowerReportStorageRecord, _
'    ByRef Records_CubeFieldOrientationProperties() As TypePowerReportStorageRecord, _
'    ByRef Records_CubeFieldExOrientationProperties() As TypePowerReportStorageRecord, _
'    ByRef Records_PivotFieldSubtotalProperties() As TypePowerReportStorageRecord, _
'    ByRef Records_PivotFieldProperties() As TypePowerReportStorageRecord, _
'    ByVal iRowNumber As Integer, ByVal sSheetHeading As String, ByVal sSheetCategory As String, _
'    ByRef vFreezeDetails As Variant, ByRef vRowRangeColWidths As Variant, _
'    ByRef vDataBodyRangeColWidth As Variant)
''Creates power pivot report from stored data
''Note this will delete any existing sheet with same name
'
'    this.Name = sSheetName
'    Set this.sht = PR_SetNewSheet(wkbReportSource, sSheetName)
'    Set this.wkb = this.sht.Parent
'    Set this.pvt = PR_CreateEmptyPivot
'
'    PR_SetPivotTableProperties Records_PivotTableProperties
'    PR_SetCubeFieldOrientation Records_CubeFieldOrientationProperties
'    PR_SetCubeFieldPropertiesExOrientation Records_CubeFieldExOrientationProperties
'    PR_SetPivotFieldSubtotals Records_PivotFieldSubtotalProperties
'    PR_SetPivotFieldProperties Records_PivotFieldProperties
'    PR_SetPivotTableRowLocation iRowNumber
'    PR_SetSheetFormat this.sht, sSheetHeading, sSheetCategory
'    PR_SetViewLayout vFreezeDetails, vRowRangeColWidths, vDataBodyRangeColWidth
'
'
'End Sub
'
'
'Public Sub CreateStandardPivotReportFromData( _
'    ByRef wkbReportSource As Workbook, _
'    ByVal sSheetName As String, _
'    ByRef Records_PivotTableProperties() As TypePowerReportStorageRecord, _
'    ByRef Records_CubeFieldOrientationProperties() As TypePowerReportStorageRecord, _
'    ByRef Records_CubeFieldExOrientationProperties() As TypePowerReportStorageRecord, _
'    ByRef Records_PivotFieldSubtotalProperties() As TypePowerReportStorageRecord, _
'    ByRef Records_PivotFieldProperties() As TypePowerReportStorageRecord, _
'    ByVal iRowNumber As Integer, ByVal sSheetHeading As String, ByVal sSheetCategory As String, _
'    ByRef vFreezeDetails As Variant, ByRef vRowRangeColWidths As Variant, _
'    ByRef vDataBodyRangeColWidth As Variant)
''Creates standard pivot report from stored data
''Note this will delete any existing sheet with same name
'
'    this.Name = sSheetName
'    Set this.sht = PR_SetNewSheet(wkbReportSource, sSheetName)
'    Set this.wkb = this.sht.Parent
'
'    '!!!!!!!!!!!!!! DO not call CreateExcelBacking - rather call underlying subs as
'    'setup is slightly different
'
''    Set This.sht = PR_SetNewSheet(wkbReportSource, sSheetName)
''    Set This.wkb = This.sht.Parent
''    Set This.pvt = PR_CreateEmptyPivot
''
''    PR_SetPivotTableProperties Records_PivotTableProperties
''    PR_SetCubeFieldOrientation Records_CubeFieldOrientationProperties
''    PR_SetCubeFieldPropertiesExOrientation Records_CubeFieldExOrientationProperties
''    PR_SetPivotFieldSubtotals Records_PivotFieldSubtotalProperties
''    PR_SetPivotFieldProperties Records_PivotFieldProperties
''    PR_SetPivotTableRowLocation iRowNumber
''    PR_SetSheetFormat sSheetHeading, sSheetCategory
''    PR_SetViewLayout vFreezeDetails, vRowRangeColWidths, vDataBodyRangeColWidth
'
'
'
'End Sub
'
'

'Sub CreateValueCopy(ByRef wkbValues As Workbook)
''Moves PowerReport as values (no pivot table) to a new workbook (and effectively)
'
'    Dim shtNew As Worksheet
'    Dim shtBacking As Worksheet
'    Dim sht As Worksheet
'    Dim sPageRangeFirstCellAddress As String
'    Dim sTable1RangeFirstCellAddress As String
'    Dim vDaxBackingQuery As Variant
'
'    'Const csDummyQueryName As String = "DummyTableForDaxQueries"
'
'
'
'    this.sht.Copy Before:=this.sht.Parent.Sheets(1)
'    Set shtNew = this.sht.Parent.Sheets(1)
'
'    shtNew.PivotTables(1).TableRange2.Clear
'
'
'    'Copy table filter range value and formats
'    'Needs to be done seperately to Table1Range for some unknwon reason as
'    'otherwise formats are not copied.
'    'https://www.contextures.com/excel-vba-pivot-table-paste-format.html
'    'On error checks in case range does not exist
'    On Error Resume Next
'    sPageRangeFirstCellAddress = this.pvt.PageRange.Cells(1).Address
'    If Err.Number = 0 Then
'        this.pvt.PageRange.Copy
'        shtNew.Range(sPageRangeFirstCellAddress).PasteSpecial xlPasteValues
'        shtNew.Range(sPageRangeFirstCellAddress).PasteSpecial xlPasteFormats
'    End If
'    On Error GoTo 0
'
'    'CopyTable1 range values and formats
'    On Error Resume Next
'    sTable1RangeFirstCellAddress = this.pvt.TableRange1.Cells(1).Address
'    If Err.Number = 0 Then
'        this.pvt.TableRange1.Copy
'        shtNew.Range(sTable1RangeFirstCellAddress).PasteSpecial xlPasteValues
'        shtNew.Range(sTable1RangeFirstCellAddress).PasteSpecial xlPasteFormats
'    End If
'    On Error GoTo 0
'
''    'Optionally create backup data to report based on a saved DAX query
''    vDaxBackingQuery = m001_Reporting_Data.PR_GetBackingQuery(This.sht.Name)
''    If Not IsNull(vDaxBackingQuery) Then
''        CreateDummyPowerPivotTable csDummyQueryName
''        Set shtBacking = PR_CreateDataQueryBackingSheet
''        PR_CreateListObjectBacking shtBacking, csDummyQueryName, This.sht.Name, vDaxBackingQuery
''    End If
''
'    'Move value report into new workbook
'    shtNew.Move Before:=wkbValues.Sheets(1)
'    wkbValues.Sheets(1).Name = this.sht.Name
'
'    wkbValues.Sheets(1).Activate
'    wkbValues.Sheets(1).Range("a1").Select
'
'End Sub


'''********************************************************************************************************
'''
'''                               Private Subs
'''
'''********************************************************************************************************
'




'Private Sub PR_AssignPivotFieldPropertiesToRecords(ByRef StorageRecords() As TypePowerReportStorageRecord, _
'    ByRef StorageIndex As Long)
''Some of Object model properties are set at pivot field level.   Others at a cubefield level
'
'
'    Dim pvtField As PivotField
'    Dim i As Long
'
'    For Each pvtField In this.sht.PivotTables(1).PivotFields
'
'        If pvtField.Name = "Values" Then
'            'Special case.  This is only case where orientation is set at a pivot
'            'field level.  Other cases are set at cube field level
'            StorageIndex = StorageIndex + 1
'            With StorageRecords(StorageIndex)
'                .SheetName = this.sht.Name
'                .Name = pvtField.Name
'                .DataType = "PivotFieldDataType"
'                .Property = "Orientation"
'                .Value = pvtField.Orientation
'                .CubeFieldPosition = ""
'            End With
'
'        ElseIf pvtField.Orientation = xlDataField Then
'
'            'NumberFormat
'            StorageIndex = StorageIndex + 1
'            With StorageRecords(StorageIndex)
'                .SheetName = this.sht.Name
'                .Name = pvtField.Name
'                .DataType = "PivotFieldDataType"
'                .Property = "NumberFormat"
'
'                'Seems to be a VBA bug requires the numberformat to be set using the pivotfield
'                'name --> does not work setting property of the pivotfield directly
'                .Value = this.sht.PivotTables(1).PivotFields(pvtField.Name).NumberFormat
'
'                .CubeFieldPosition = ""
'            End With
'
'        Else
'
'            'LayoutBlankLine
'            StorageIndex = StorageIndex + 1
'            With StorageRecords(StorageIndex)
'                .SheetName = this.sht.Name
'                .Name = pvtField.Name
'                .DataType = "PivotFieldDataType"
'                .Property = "LayoutBlankLine"
'                .Value = pvtField.LayoutBlankLine
'                .CubeFieldPosition = ""
'            End With
'
'            'LayoutCompactRow
'            StorageIndex = StorageIndex + 1
'            With StorageRecords(StorageIndex)
'                .SheetName = this.sht.Name
'                .Name = pvtField.Name
'                .DataType = "PivotFieldDataType"
'                .Property = "LayoutCompactRow"
'                .Value = pvtField.LayoutCompactRow
'                .CubeFieldPosition = ""
'            End With
'
'            'LayoutForm
'            StorageIndex = StorageIndex + 1
'            With StorageRecords(StorageIndex)
'                .SheetName = this.sht.Name
'                .Name = pvtField.Name
'                .DataType = "PivotFieldDataType"
'                .Property = "LayoutForm"
'                .Value = pvtField.LayoutForm
'                .CubeFieldPosition = ""
'            End With
'
'            'LayoutPageBreak
'            StorageIndex = StorageIndex + 1
'            With StorageRecords(StorageIndex)
'                .SheetName = this.sht.Name
'                .Name = pvtField.Name
'                .DataType = "PivotFieldDataType"
'                .Property = "LayoutPageBreak"
'                .Value = pvtField.LayoutPageBreak
'                .CubeFieldPosition = ""
'            End With
'
'            'RepeatLables
'            StorageIndex = StorageIndex + 1
'            With StorageRecords(StorageIndex)
'                .SheetName = this.sht.Name
'                .Name = pvtField.Name
'                .DataType = "PivotFieldDataType"
'                .Property = "RepeatLabels"
'                .Value = pvtField.RepeatLabels
'                .CubeFieldPosition = ""
'            End With
'
'
'            'Subotal(note the property is indexed)
'            StorageIndex = StorageIndex + 1
'            With StorageRecords(StorageIndex)
'                .SheetName = this.sht.Name
'                .Name = pvtField.Name
'                .DataType = "PivotFieldDataType"
'                .Property = "Subtotals"
'                'Note the index one here to designate automatic for powerpivot
'                .Value = pvtField.Subtotals(1)
'                .CubeFieldPosition = ""
'            End With
'
'            'Only save below if field has subtotals:
'            If pvtField.Subtotals(1) Then
'
'                'SubtotalName
'                StorageIndex = StorageIndex + 1
'                With StorageRecords(StorageIndex)
'                    .SheetName = this.sht.Name
'                    .Name = pvtField.Name
'                    .DataType = "PivotFieldDataType"
'                    .Property = "SubtotalName"
'                    .Value = pvtField.SubtotalName
'                    .CubeFieldPosition = ""
'                End With
'
'                'LayoutSubtotalLocation
'                StorageIndex = StorageIndex + 1
'                With StorageRecords(StorageIndex)
'                    .SheetName = this.sht.Name
'                    .Name = pvtField.Name
'                    .DataType = "PivotFieldDataType"
'                    .Property = "LayoutSubtotalLocation"
'                    .Value = pvtField.LayoutSubtotalLocation
'                    .CubeFieldPosition = ""
'                End With
'
'            End If
'
'
'        End If
'
'    Next pvtField
'
'End Sub
'
'
'Private Sub PR_AssignViewLayoutToRecords(ByRef StorageRecords() As TypePowerReportStorageRecord, _
'    ByRef StorageIndex As Long)
'' Reads freeze panes and col widths.  This is performed last to mirror the creation as this can only
''be created last once the report is in place
'
'    Dim rng As Range
'    Dim sRowRangeColWidthsStr As String
'    Dim sDataRangeColWidthsStr As String
'
'    'Required to activate sheet to read the pane freeze location
'    this.sht.Activate
'
'    'Read pane freeze if applicable into a pipe delimited format
'    If (ActiveWindow.SplitColumn <> 0 And CStr(ActiveWindow.SplitColumn) <> "") Or _
'        (ActiveWindow.SplitRow <> 0 And CStr(ActiveWindow.SplitRow) <> "") Then
'
'        StorageIndex = StorageIndex + 1
'        With StorageRecords(StorageIndex)
'            .SheetName = this.sht.Name
'            .Name = this.sht.Name
'            .DataType = "ViewLayoutDataType"
'            .Property = "FreezePanes"
'            .Value = ActiveWindow.SplitRow & "|" & ActiveWindow.SplitColumn
'        End With
'
'    End If
'
'    'Save pivot row column widths as pipe delimited string
'    If PR_PivotTableHasRowRange Then
'        StorageIndex = StorageIndex + 1
'        sRowRangeColWidthsStr = ""
'        For Each rng In this.sht.PivotTables(1).RowRange.Rows(1).Cells
'            If sRowRangeColWidthsStr = "" Then
'                sRowRangeColWidthsStr = CStr(rng.ColumnWidth)
'            Else
'                sRowRangeColWidthsStr = sRowRangeColWidthsStr & "|" & CStr(rng.ColumnWidth)
'            End If
'        Next rng
'        With StorageRecords(StorageIndex)
'            .SheetName = this.sht.Name
'            .Name = this.sht.Name
'            .DataType = "ViewLayoutDataType"
'            .Property = "PivotRowRangeColWidths"
'            .Value = sRowRangeColWidthsStr
'        End With
'    End If
'
'    'Save data column  - all data columns are set at same width which is obtained from
'    'first column
'    If PR_PivotTableHasRowRange Then
'        StorageIndex = StorageIndex + 1
'        sDataRangeColWidthsStr = CStr(this.sht.PivotTables(1).DataBodyRange.Cells(1).ColumnWidth)
'        With StorageRecords(StorageIndex)
'            .SheetName = this.sht.Name
'            .Name = this.sht.Name
'            .DataType = "ViewLayoutDataType"
'            .Property = "PivotDataBodyRangeColWidths"
'            .Value = sDataRangeColWidthsStr
'        End With
'    End If
'
'End Sub
'
'
'
'Private Function PR_SetNewSheet(ByRef wkbReportSource As Workbook, ByVal sSheetName)
''Note this will delete any existing sSheetName if it exists
'
'    Dim sht As Worksheet
'
'    On Error Resume Next
'    wkbReportSource.Sheets(sSheetName).Delete
'    On Error GoTo 0
'    Set sht = wkbReportSource.Sheets.Add(After:=wkbReportSource.Sheets(wkbReportSource.Sheets.Count))
'    sht.Name = sSheetName
'
'    Set PR_SetNewSheet = sht
'
'End Function
'
'
'Private Function PR_CreateEmptyPivot() As PivotTable
'
'
'    'Create empty pivot table from data model
'    'Create in 1st row then shfit down to allow for number of filter fields that are
'    'unknown at the moment
'    Set PR_CreateEmptyPivot = this.wkb.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
'        this.wkb.Connections("ThisWorkbookDataModel"), Version:=6). _
'        CreatePivotTable(this.sht.Range("B1"))
'
'End Function
'
'
'Private Sub PR_SetPivotTableProperties(ByRef Records() As TypePowerReportStorageRecord)
'
'    Dim i As Long
'
'    For i = LBound(Records) To UBound(Records)
'        CallByName this.pvt, Records(i).Property, VbLet, Records(i).Value
'    Next i
'
'End Sub
'
'
'
'Private Sub PR_SetCubeFieldOrientation(ByRef Records() As TypePowerReportStorageRecord)
''This property brings fields onto pivot table.  Needs to be set before other properties
'
'
'    Dim i As Long
'    Dim sCubeName As String
'    Dim lCubeOrientation As Long
'
'    For i = LBound(Records) To UBound(Records)
'        sCubeName = Records(i).Name
'        lCubeOrientation = Records(i).Value
'        this.pvt.CubeFields(sCubeName).Orientation = lCubeOrientation
'    Next i
'
'End Sub
'
'
'Private Sub PR_SetCubeFieldPropertiesExOrientation(ByRef Records() As TypePowerReportStorageRecord)
''Orientation property needs to be set before this sub is called
'
'
'    Dim i As Long
'    Dim sCubeName As String
'    Dim sCubeProperty As String
'    Dim vCubeValue As Variant
'
'    For i = LBound(Records) To UBound(Records)
'        sCubeName = Records(i).Name
'        sCubeProperty = Records(i).Property
'        vCubeValue = Records(i).Value
'        CallByName this.pvt.CubeFields(sCubeName), sCubeProperty, VbLet, vCubeValue
'    Next i
'
'End Sub
'
'
'Private Sub PR_SetPivotFieldSubtotals(ByRef Records() As TypePowerReportStorageRecord)
''Needs to be done first as it has flow on impacts to other properties.
''Also slightly different as the subtotal property requires an index
'
'    Dim i As Long
'    Dim sPivotFieldName As String
'    Dim sPivotFieldProperty As String
'    Dim vPivotFieldValue As Variant
'
'    For i = LBound(Records) To UBound(Records)
'        sPivotFieldName = Records(i).Name
'        sPivotFieldProperty = Records(i).Property
'        vPivotFieldValue = Records(i).Value
'        this.pvt.PivotFields(sPivotFieldName).Subtotals(1) = vPivotFieldValue
'    Next i
'
'End Sub
'
'
'Private Sub PR_SetPivotFieldProperties(ByRef Records() As TypePowerReportStorageRecord)
'
'    Dim i As Long
'    Dim sPivotFieldName As String
'    Dim sPivotFieldProperty As String
'    Dim vPivotFieldValue As Variant
'
'    For i = LBound(Records) To UBound(Records)
'        sPivotFieldName = Records(i).Name
'        sPivotFieldProperty = Records(i).Property
'        vPivotFieldValue = Records(i).Value
'        CallByName this.pvt.PivotFields(sPivotFieldName), sPivotFieldProperty, VbLet, vPivotFieldValue
'    Next i
'
'End Sub
'
'
'
'
'Sub PR_SetPivotTableRowLocation(ByVal iRowNumber As Integer)
''Move pivot table down now that filters are in place
'
'    Dim sRangeStr As String
'
'    sRangeStr = "1:" & CLng(iRowNumber - 1)
'    this.sht.Range(sRangeStr).Insert xlShiftDown
'
'End Sub
'
'
'Private Sub PR_SetSheetFormat(ByRef sht As Worksheet, ByVal sSheetHeading As String, _
'    ByVal sSheetCategory As String)
''Sets sht format includiing range names
'
'    With sht
'
'        'Necessary to activate to set gridlines and zoom
'        this.wkb.Activate
'        .Activate
'        ActiveWindow.DisplayGridlines = False
'        ActiveWindow.Zoom = 80
'
'        .Cells.Font.Name = "Calibri"
'        .Cells.Font.Size = 11
'        .DisplayPageBreaks = False
'        .Columns("A:A").ColumnWidth = 4
'
'        'Add sheet heading named range and format
'        .Names.Add Name:="SheetHeading", RefersTo:="=$B$2"
'        .Range("SheetHeading") = sSheetHeading
'        .Range("SheetHeading").Font.Bold = True
'        .Range("SheetHeading").Font.Size = 16
'
'        'Add sheet category named range and format
'        .Names.Add Name:="SheetCategory", RefersTo:="=$A$1"
'        .Range("SheetCategory") = sSheetCategory
'        .Range("SheetCategory").Font.Color = RGB(170, 170, 170)
'        .Range("SheetCategory").Font.Size = 8
'
'    End With
'
'End Sub
'
'
'
'
'
'Private Sub PR_SetViewLayout(ByRef vFreezeDetails As Variant, ByRef vRowRangeColWidths As Variant, _
'    ByRef vDataBodyRangeColWidth As Variant)
''Freezes panes and sets the column widths.  Can only be set once the rest of the pivot table is
''in place
'
'    Dim iFreezeRow As Integer
'    Dim iFreezeCol As Integer
'    Dim vArrayRowRangeColWidths As Variant
'    Dim i As Integer
'
'    'Set freeze pane locations
'    If Not IsNull(vFreezeDetails) Then
'        iFreezeRow = Split(vFreezeDetails, "|")(0)
'        iFreezeCol = Split(vFreezeDetails, "|")(1)
'        this.sht.Activate
'        this.sht.Cells(iFreezeRow + 1, iFreezeCol + 1).Select
'        ActiveWindow.FreezePanes = True
'    End If
'
'    'Set Row Range column widths
'    If Not IsNull(vRowRangeColWidths) Then
'        vArrayRowRangeColWidths = Split(vRowRangeColWidths, "|")
'        For i = LBound(vArrayRowRangeColWidths) To UBound(vArrayRowRangeColWidths)
'            this.sht.PivotTables(1).RowRange.Rows(1).Cells(i + 1).ColumnWidth = _
'                vArrayRowRangeColWidths(i)
'        Next i
'    End If
'
'    'Set data body range column widths (all widths are set identical sizes)
'    If Not IsNull(vDataBodyRangeColWidth) Then
'        this.sht.PivotTables(1).DataBodyRange.EntireColumn.ColumnWidth = vDataBodyRangeColWidth
'    End If
'
'
'End Sub
'
'
'

'
'
'Private Sub GetDaxFileItemsInFolder(ByRef FileItems() As Scripting.File)
''Returns an array of files (which can be used to get filename, path etc)
''(Cannot create function due to recursive nature of the code)
'
'    Dim sFolderPath As String
'    Dim fso As Scripting.FileSystemObject
'    Dim SourceFolder As Scripting.Folder
'    Dim SubFolder As Scripting.Folder
'    Dim FileItem As Scripting.File
'    Dim bIsFirstItemInArray As Boolean
'
'    sFolderPath = this.wkb.Path & Application.PathSeparator & csDaxQuerySubDir
'    Set fso = New Scripting.FileSystemObject
'    Set SourceFolder = fso.GetFolder(sFolderPath)
'    bIsFirstItemInArray = True
'
'    For Each FileItem In SourceFolder.Files
'
'        If bIsFirstItemInArray Then
'            bIsFirstItemInArray = False
'            ReDim FileItems(0)
'        Else
'            ReDim Preserve FileItems(UBound(FileItems) + 1)
'        End If
'
'        Set FileItems(UBound(FileItems)) = FileItem
'
'    Next FileItem
'
'
'    Set FileItem = Nothing
'    Set SourceFolder = Nothing
'    Set fso = Nothing
'
'
'End Sub
'
'Private Function CreateDataBackingSheet(ByRef wkbSource As Workbook, ByVal sDataSheetName As String) As Worksheet
''Creates a backing data sheet generated via dax query (which is later converted to values) +
''spreadsheet listobject styls formulas
'
'    Dim sht As Worksheet
'
'    On Error Resume Next
'    wkbSource.Sheets(sDataSheetName).Delete
'    On Error GoTo 0
'
'    Set sht = wkbSource.Worksheets.Add(After:=wkbSource.Worksheets(wkbSource.Worksheets.Count))
'    sht.Name = sDataSheetName
'    Set CreateDataBackingSheet = sht
'
'End Function
'
'Private Sub CreateEmptyPowerPivotTable()
''Creates empty table loaded into Power Pivot to form source of listobject connection for DAX queries
'
'    Dim sQueryStr As String
'    Dim sTestName As String
'    Dim bEmptyTableExists As Boolean
'
'
'    On Error Resume Next
'        sTestName = this.wkb.Model.ModelTables(csEmptyDaxTableName).Name
'        If Err.Number = 0 Then
'            bEmptyTableExists = True
'        Else
'            bEmptyTableExists = False
'        End If
'    On Error GoTo 0
'
'
'    If Not bEmptyTableExists Then
'        sQueryStr = "let" & vbCr & _
'            "    tbl = Table.FromRecords({[NullHeader = """"]}), " & vbCr & _
'            "    ChangedType = Table.TransformColumnTypes(tbl, {{""NullHeader"", type text}})" & vbCr & _
'            "in" & vbCr & _
'            "    ChangedType"
'
'        this.wkb.Queries.Add Name:=csEmptyDaxTableName, Formula:=sQueryStr
'        this.wkb.Connections.Add2 _
'            Name:=csEmptyDaxTableName, _
'            Description:="Connection to the 'EmptyTable' query in the workbook.", _
'            ConnectionString:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & _
'                csEmptyDaxTableName & ";Extended Properties=", _
'            CommandText:="" & csEmptyDaxTableName & "", _
'            lCmdtype:=6, _
'            CreateModelConnection:=True, _
'            ImportRelationships:=False
'
'    End If
'
'End Sub
'
'
'
'Private Sub CreateDaxQueryListObject()
''Utilsied for creating optional backing tab of data when values report sheet is created.
''This sub creates the ListObject query.
'
'    Dim iFileNo As Integer
'    Dim sDaxQuery As String
'    Dim fso As FileSystemObject
'    Dim sQueryFilePath As String
'
'    sQueryFilePath = this.wkb.Path & Application.PathSeparator & csDaxQuerySubDir & _
'        Application.PathSeparator & this.Name & ".dax"
'
'    'Read DAX query from file
'    iFileNo = FreeFile
'    Open sQueryFilePath For Input As #iFileNo
'    sDaxQuery = Input$(LOF(iFileNo), iFileNo)
'    Close #iFileNo
'
'    'If query is directly saved from DAX Studio it seems to save unicode format. Only
'    'issues are symbols at start hence the replace below.  ADO can be used if more
'    'issues arrise in the future.
'    'https://stackoverflow.com/questions/13851473/read-utf-8-text-file-in-vbscript/13855268#13855268
'    sDaxQuery = Replace(sDaxQuery, "﻿", "")
'
'    Set this.loDataBacking = this.shtDataBacking.ListObjects.Add( _
'        SourceType:=xlSrcModel, _
'        Source:=this.wkb.Connections(csEmptyDaxTableName), _
'        Destination:=this.shtDataBacking.Range("$B$6"))
'
'
'    With this.loDataBacking.TableObject
'        .PreserveFormatting = True
'        .RefreshStyle = xlOverwriteCells
'        .AdjustColumnWidth = True
'        .ListObject.Name = "tbl_" & this.shtDataBacking.Name
'    End With
'
'    With this.loDataBacking.TableObject.WorkbookConnection.OLEDBConnection
'        .CommandText = Array(sDaxQuery)
'        .CommandType = xlCmdDAX
'    End With
'
'    this.loDataBacking.Refresh
'
'End Sub
'
'
'
'
'Private Sub PR_FormatListObj()
'
'    Dim sty As TableStyle
'
'    On Error Resume Next
'    this.wkb.TableStyles.Add ("SpreadsheetBiStyle")
'    On Error GoTo 0
'    Set sty = this.wkb.TableStyles("SpreadsheetBiStyle")
'
'    'Set Header Format
'    With sty.TableStyleElements(xlHeaderRow)
'        .Interior.Color = RGB(68, 114, 196)
'        .Font.Color = RGB(255, 255, 255)
'        .Font.Bold = True
'        .Borders.item(xlEdgeTop).LineStyle = xlSolid
'        .Borders.item(xlEdgeTop).Weight = xlMedium
'        .Borders.item(xlEdgeBottom).LineStyle = xlSolid
'        .Borders.item(xlEdgeBottom).Weight = xlMedium
'    End With
'
'    'Set row stripe format
'    sty.TableStyleElements(xlRowStripe1).Interior.Color = RGB(217, 217, 217)
'    sty.TableStyleElements(xlRowStripe2).Interior.Color = RGB(255, 255, 255)
'
'    'Set whole table bottom edge format
'    sty.TableStyleElements(xlWholeTable).Borders.item(xlEdgeBottom).LineStyle = xlSolid
'    sty.TableStyleElements(xlWholeTable).Borders.item(xlEdgeBottom).Weight = xlMedium
'
'
'    'Apply custom style and set other attributes
'    this.loDataBacking.TableStyle = "SpreadsheetBiStyle"
'    With this.loDataBacking.HeaderRowRange
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlTop
'        .WrapText = True
'        .Orientation = 0
'    End With
'
'    If Not this.loDataBacking.DataBodyRange Is Nothing Then
'        this.loDataBacking.DataBodyRange.EntireColumn.AutoFit
'    End If
'
'    'Set DAX query range font as blue (this takes place before any formula columns are added
'    On Error Resume Next 'In event no data returned and then no databodyrange
'    this.loDataBacking.DataBodyRange.Font.Color = RGB(0, 112, 192)
'    On Error GoTo 0
'
'
'End Sub
'
'
'
'
'Private Sub ConvertDaxQueryBackingToValues()
'
'    this.loDataBacking.TableObject.Delete
'
'End Sub
'
'

