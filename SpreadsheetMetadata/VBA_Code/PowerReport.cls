VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PowerReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Use "This" declaration as an easy way to get intellisense to the classes private variables
'https://rubberduckvba.wordpress.com/2020/02/27/vba-classes-gateway-to-solid/
Private Type TypePowerReport
    sht As Worksheet
    StorageRecords() As TypePowerReportStorageRecord   'Utilised for storage read write
    StorageIndex As Long                               'Tracks storage record index
End Type
    
Private This As TypePowerReport


Public Function AssignToExistingSheet(ByRef shtToAssign As Worksheet) As Boolean

    Dim bValidSheet As Boolean

    bValidSheet = True

    'Check if sht meets criteria for a Power Pivot Report
    On Error Resume Next
    Select Case True
    Case shtToAssign.PivotTables.Count <> 1
        bValidSheet = False
    Case shtToAssign.PivotTables(1).PivotCache.WorkbookConnection.Name <> "ThisWorkbookDataModel"
        bValidSheet = False
    Case IsError(Evaluate("='" & shtToAssign.Name & "'!SheetCategory"))
        bValidSheet = False
    Case IsError(Evaluate("='" & shtToAssign.Name & "'!SheetHeading"))
        bValidSheet = False
    Case Err.Number <> 0
        bValidSheet = False
    End Select

    On Error GoTo 0



    If bValidSheet Then
        Set This.sht = shtToAssign
        AssignToExistingSheet = True
    Else
        AssignToExistingSheet = False
    End If
    
    
End Function

Public Property Get SheetName() As String
    SheetName = This.sht.Name
End Property


Public Property Get HasFilters() As Boolean
'Returns true if any of the Cubefields are filtered

    Dim pvtCubeField As cubeField
        
    HasFilters = False
     
    For Each pvtCubeField In This.sht.PivotTables(1).CubeFields
        If Not pvtCubeField.AllItemsVisible Then
            HasFilters = True
        End If
    Next pvtCubeField

End Property


Public Sub SavePowerReportStructure()
    
    m001_DataAccess.PR_SetupOrAssignStorage
    m001_DataAccess.PR_DeleteExistingData This.sht.Name
    ReDim This.StorageRecords(cPR_MaxStorageRecords)
    
    This.StorageIndex = 0
    PR_GetReportSheetProperties
    PR_GetPivotTableProperties
    PR_GetCubeFieldProperties
    PR_GetPivotFieldProperties
    PR_GetViewLayout
    ReDim Preserve This.StorageRecords(This.StorageIndex)
    WritePowerReportPropertiesToStorage
    
End Sub

Public Sub CreateFromStoredData(ByRef wkb As Workbook, ByVal sSheetName)
'Note this will delete any existing sheet with same name
'sSheetName needs to already exist in stored data
    
    Dim pvt As PivotTable
    
    Set This.sht = PR_SetSheet(wkb, sSheetName)
    PR_CreateEmptyPivot
    PR_SetPivotTableProperties
    PR_SetCubeFieldOrientation
    PR_SetCubeFieldPropertiesExOrientation
    PR_SetPivotFieldSubtotals
    PR_SetPivotFieldProperties
    PR_SetPivotTableLocation
    PR_SetSheetFormat
    PR_SetViewLayout


End Sub


Sub CreateValueReport()
'Create a copy of PowerReport as values (no pivot table) in a new workbook

    Dim pvt As PivotTable
    Dim shtNew As Worksheet
    Dim shtBacking As Worksheet
    Dim sht As Worksheet
    Dim sPageRangeFirstCellAddress As String
    Dim sTable1RangeFirstCellAddress As String
    Dim vDaxBackingQuery As Variant
    Dim wkbValues As Workbook
    
    Const csDummyQueryName As String = "DummyTableForDaxQueries"

    Set pvt = This.sht.PivotTables(1)

    This.sht.Copy Before:=This.sht.Parent.Sheets(1)
    Set shtNew = This.sht.Parent.Sheets(1)
    
    shtNew.PivotTables(1).TableRange2.Clear
    
    
    'Copy table filter range value and formats
    'Needs to be done seperately to Table1Range for some unknwon reason as
    'otherwise formats are not copied.
    'https://www.contextures.com/excel-vba-pivot-table-paste-format.html
    'On error checks in case range does not exist
    On Error Resume Next
    sPageRangeFirstCellAddress = pvt.PageRange.Cells(1).Address
    If Err.Number = 0 Then
        pvt.PageRange.Copy
        shtNew.Range(sPageRangeFirstCellAddress).PasteSpecial xlPasteValues
        shtNew.Range(sPageRangeFirstCellAddress).PasteSpecial xlPasteFormats
    End If
    On Error GoTo 0
    
    'CopyTable1 range values and formats
    On Error Resume Next
    sTable1RangeFirstCellAddress = pvt.TableRange1.Cells(1).Address
    If Err.Number = 0 Then
        pvt.TableRange1.Copy
        shtNew.Range(sTable1RangeFirstCellAddress).PasteSpecial xlPasteValues
        shtNew.Range(sTable1RangeFirstCellAddress).PasteSpecial xlPasteFormats
    End If
    On Error GoTo 0
        
    'Optionally create backup data to report based on a saved DAX query
    vDaxBackingQuery = m001_DataAccess.PR_GetBackingQuery(This.sht.Name)
    If Not IsNull(vDaxBackingQuery) Then
        CreateDummyPowerPivotTable csDummyQueryName
        Set shtBacking = PR_CreateDataQueryBackingSheet
        PR_CreateListObjectBacking shtBacking, csDummyQueryName, This.sht.Name, vDaxBackingQuery
    End If
    
    'Move value report into new workbook
    Set wkbValues = Workbooks.Add
    shtNew.Move Before:=wkbValues.Sheets(1)
    wkbValues.Sheets(1).Name = This.sht.Name
    
    'If applicable move backing sheet into the above created new workbook
    If Not IsNull(vDaxBackingQuery) Then
        shtBacking.Move After:=wkbValues.Sheets(1)
    End If

    For Each sht In wkbValues.Sheets
        If sht.Name <> This.sht.Name And sht.Name <> This.sht.Name & "_Backing" Then
            sht.Delete
        End If
    Next sht
        
    wkbValues.Sheets(1).Activate
    wkbValues.Sheets(1).Range("a1").Select

End Sub

'********************************************************************************************************
'
'                               Private Subs
'
'********************************************************************************************************


Private Sub PR_GetReportSheetProperties()
'Reads sheet properties into records for later saving

    'Read Sheet Heading
    With This.StorageRecords(This.StorageIndex)
        .SheetName = This.sht.Name
        .Name = This.sht.Name
        .DataType = "SheetDataType"
        .Property = "SheetHeading"
        .Value = This.sht.Range("SheetHeading")
    End With
    
    'Read Sheet Category
    This.StorageIndex = This.StorageIndex + 1
    With This.StorageRecords(This.StorageIndex)
        .SheetName = This.sht.Name
        .Name = This.sht.Name
        .DataType = "SheetDataType"
        .Property = "SheetCategory"
        .Value = This.sht.Range("SheetCategory")
    End With
    
    
End Sub


Private Sub PR_GetPivotTableProperties()
'Reads pivot table properties into records for later saving

    Dim Properties As Variant
    Dim item As Variant
    Dim i As Long

    Properties = Array("LayoutRowDefault", "PageFieldWrapCount", "CompactRowIndent", _
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
        
        For Each item In Properties
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.sht.Name
                .Name = This.sht.PivotTables(1).Name
                .DataType = "PivotTableDataType"
                .Property = item
                .Value = CallByName(This.sht.PivotTables(1), item, VbGet)
                .CubeFieldPosition = ""
            End With
        Next item
    
End Sub

Private Sub PR_GetCubeFieldProperties()
'Some of Object model properties are set at cube field level.   Others at a PivotField level
'Simpler not to loop a Properties array due to exceptions that need to be handled as below.
 
    Dim pvtCubeField As cubeField
    'Dim Properties As Variant
    Dim i As Long

    'Properties = Array("Caption", "Orientation", "Position")

    For Each pvtCubeField In This.sht.PivotTables(1).CubeFields
        If pvtCubeField.Orientation <> xlHidden Then
            
                'Capture position along with the cubefield orientation.  For some reason it
                'doesnt work setting position in VBA, therefore when data is read it is sorted
                'by position for purposes of creating pivot table
                This.StorageIndex = This.StorageIndex + 1
                With This.StorageRecords(This.StorageIndex)
                    .SheetName = This.sht.Name
                    .Name = pvtCubeField.Name
                    .DataType = "PivotCubeFieldDataType"
                    .Property = "Orientation"
                    .Value = pvtCubeField.Orientation
                    .CubeFieldPosition = pvtCubeField.Position
                End With
            
                This.StorageIndex = This.StorageIndex + 1
                With This.StorageRecords(This.StorageIndex)
                    .SheetName = This.sht.Name
                    .Name = pvtCubeField.Name
                    .DataType = "PivotCubeFieldDataType"
                    .Property = "Caption"
                    .Value = pvtCubeField.Caption
                    .CubeFieldPosition = ""
                End With
            
        End If
    Next pvtCubeField


End Sub


Private Sub PR_GetPivotFieldProperties()
'Some of Object model properties are set at pivot field level.   Others at a cubefield level


    Dim pvtField As PivotField
    Dim i As Long

    For Each pvtField In This.sht.PivotTables(1).PivotFields

        If pvtField.Name = "Values" Then
            'Special case.  This is only case where orientation is set at a pivot
            'field level.  Other cases are set at cube field level
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.sht.Name
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "Orientation"
                .Value = pvtField.Orientation
                .CubeFieldPosition = ""
            End With

        ElseIf pvtField.Orientation = xlDataField Then

            'NumberFormat
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.sht.Name
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "NumberFormat"
                
                'Seems to be a VBA bug requires the numberformat to be set using the pivotfield
                'name --> does not work setting property of the pivotfield directly
                .Value = This.sht.PivotTables(1).PivotFields(pvtField.Name).NumberFormat
                
                .CubeFieldPosition = ""
            End With

        Else
        
            'LayoutBlankLine
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.sht.Name
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "LayoutBlankLine"
                .Value = pvtField.LayoutBlankLine
                .CubeFieldPosition = ""
            End With
            
            'LayoutCompactRow
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.sht.Name
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "LayoutCompactRow"
                .Value = pvtField.LayoutCompactRow
                .CubeFieldPosition = ""
            End With
            
            'LayoutForm
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.sht.Name
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "LayoutForm"
                .Value = pvtField.LayoutForm
                .CubeFieldPosition = ""
            End With
            
            'LayoutPageBreak
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.sht.Name
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "LayoutPageBreak"
                .Value = pvtField.LayoutPageBreak
                .CubeFieldPosition = ""
            End With
            
            'RepeatLables
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.sht.Name
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "RepeatLabels"
                .Value = pvtField.RepeatLabels
                .CubeFieldPosition = ""
            End With
            
            
            'Subotal(note the property is indexed)
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.sht.Name
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "Subtotals"
                'Note the index one here to designate automatic for powerpivot
                .Value = pvtField.Subtotals(1)
                .CubeFieldPosition = ""
            End With
            
            'Only save below if field has subtotals:
            If pvtField.Subtotals(1) Then
                
                'SubtotalName
                This.StorageIndex = This.StorageIndex + 1
                With This.StorageRecords(This.StorageIndex)
                    .SheetName = This.sht.Name
                    .Name = pvtField.Name
                    .DataType = "PivotFieldDataType"
                    .Property = "SubtotalName"
                    .Value = pvtField.SubtotalName
                    .CubeFieldPosition = ""
                End With
                
                'LayoutSubtotalLocation
                This.StorageIndex = This.StorageIndex + 1
                With This.StorageRecords(This.StorageIndex)
                    .SheetName = This.sht.Name
                    .Name = pvtField.Name
                    .DataType = "PivotFieldDataType"
                    .Property = "LayoutSubtotalLocation"
                    .Value = pvtField.LayoutSubtotalLocation
                    .CubeFieldPosition = ""
                End With
                
            End If
                    
            
        End If

    Next pvtField

End Sub


Private Sub PR_GetViewLayout()
' Reads freeze panes and col widths.  This is performed last to mirror the creation as this can only
'be created last once the report is in place

    Dim rng As Range
    Dim sRowRangeColWidthsStr As String
    Dim sDataRangeColWidthsStr As String

    'Required to activate sheet to read the pane freeze location
    This.sht.Activate
    
    'Read pane freeze if applicable into a pipe delimited format
    If (ActiveWindow.SplitColumn <> 0 And CStr(ActiveWindow.SplitColumn) <> "") Or _
        (ActiveWindow.SplitRow <> 0 And CStr(ActiveWindow.SplitRow) <> "") Then
        
        This.StorageIndex = This.StorageIndex + 1
        With This.StorageRecords(This.StorageIndex)
            .SheetName = This.sht.Name
            .Name = This.sht.Name
            .DataType = "ViewLayoutDataType"
            .Property = "FreezePanes"
            .Value = ActiveWindow.SplitRow & "|" & ActiveWindow.SplitColumn
        End With
    
    End If
    
    'Save pivot row column widths as pipe delimited string
    If PR_PivotTableHasRowRange Then
        This.StorageIndex = This.StorageIndex + 1
        sRowRangeColWidthsStr = ""
        For Each rng In This.sht.PivotTables(1).RowRange.Rows(1).Cells
            If sRowRangeColWidthsStr = "" Then
                sRowRangeColWidthsStr = CStr(rng.ColumnWidth)
            Else
                sRowRangeColWidthsStr = sRowRangeColWidthsStr & "|" & CStr(rng.ColumnWidth)
            End If
        Next rng
        With This.StorageRecords(This.StorageIndex)
            .SheetName = This.sht.Name
            .Name = This.sht.Name
            .DataType = "ViewLayoutDataType"
            .Property = "PivotRowRangeColWidths"
            .Value = sRowRangeColWidthsStr
        End With
    End If

    'Save data column  - all data columns are set at same width which is obtained from
    'first column
    If PR_PivotTableHasRowRange Then
        This.StorageIndex = This.StorageIndex + 1
        sDataRangeColWidthsStr = CStr(This.sht.PivotTables(1).DataBodyRange.Cells(1).ColumnWidth)
        With This.StorageRecords(This.StorageIndex)
            .SheetName = This.sht.Name
            .Name = This.sht.Name
            .DataType = "ViewLayoutDataType"
            .Property = "PivotDataBodyRangeColWidths"
            .Value = sDataRangeColWidthsStr
        End With
    End If

End Sub



Private Sub WritePowerReportPropertiesToStorage()

    Dim i As Long

    For i = LBound(This.StorageRecords) To UBound(This.StorageRecords)
        m001_DataAccess.PR_WriteRecords _
            This.StorageRecords(i).SheetName, This.StorageRecords(i).Name, _
            This.StorageRecords(i).DataType, This.StorageRecords(i).Property, _
            This.StorageRecords(i).Value, This.StorageRecords(i).CubeFieldPosition
    Next i

End Sub


Private Function PR_SetSheet(ByRef wkb As Workbook, ByVal sSheetName)
'Note this will delete any existing sSheetName if it exists

    Dim sht As Worksheet

    On Error Resume Next
    wkb.Sheets(sSheetName).Delete
    On Error GoTo 0
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    sht.Name = sSheetName
    
    Set PR_SetSheet = sht

End Function

Private Function PR_CreateEmptyPivot()

    Dim wkb As Workbook
    
    Set wkb = This.sht.Parent

    'Create empty pivot table from data model
    'Create in 1st row then shfit down to allow for number of filter fields that are
    'unknown at the moment
    Set PR_CreateEmptyPivot = wkb.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
        ActiveWorkbook.Connections("ThisWorkbookDataModel"), Version:=6). _
        CreatePivotTable(This.sht.Range("B1"))

End Function

Private Sub PR_SetPivotTableProperties()

    Dim pvt As PivotTable
    Dim i As Long
    
    Set pvt = This.sht.PivotTables(1)

    'Read the Orientation field of the pivot cubes into StorageRecords array
    m001_DataAccess.PR_GetPivotTableProperties _
        This.sht.Name, This.StorageRecords
    
    For i = LBound(This.StorageRecords) To UBound(This.StorageRecords)
        CallByName pvt, This.StorageRecords(i).Property, VbLet, This.StorageRecords(i).Value
    Next i

End Sub

Private Sub PR_SetCubeFieldOrientation()
'This property brings fields onto pivot table.  Needs to be set before other properties

    Dim pvt As PivotTable
    Dim i As Long
    Dim sCubeName As String
    Dim lCubeOrientation As Long
    
    Set pvt = This.sht.PivotTables(1)

    'Read the Orientation field of the pivot cubes into StorageRecords array
    m001_DataAccess.PR_GetPivotCubeFieldDataOrientationSortedByCubeFieldPosition _
        This.sht.Name, This.StorageRecords
    
    For i = LBound(This.StorageRecords) To UBound(This.StorageRecords)
        sCubeName = This.StorageRecords(i).Name
        lCubeOrientation = This.StorageRecords(i).Value
        pvt.CubeFields(sCubeName).Orientation = lCubeOrientation
    Next i

End Sub


Private Sub PR_SetCubeFieldPropertiesExOrientation()
'Orientation property needs to be set before this sub is called


    Dim pvt As PivotTable
    Dim i As Long
    Dim sCubeName As String
    Dim sCubeProperty As String
    Dim vCubeValue As Variant
    
    Set pvt = This.sht.PivotTables(1)

    'Read the Orientation field of the pivot cubes into StorageRecords array
    m001_DataAccess.PR_GetPivotCubeFieldDataPropertiesExOrientation This.sht.Name, This.StorageRecords
    
    For i = LBound(This.StorageRecords) To UBound(This.StorageRecords)
        sCubeName = This.StorageRecords(i).Name
        sCubeProperty = This.StorageRecords(i).Property
        vCubeValue = This.StorageRecords(i).Value
        CallByName pvt.CubeFields(sCubeName), sCubeProperty, VbLet, vCubeValue
    Next i

End Sub

Private Sub PR_SetPivotFieldSubtotals()
'Needs to be done first as it has flow on impacts to other properties.
'Also slightly different as the subtotal property requires an index

    Dim pvt As PivotTable
    Dim i As Long
    Dim sPivotFieldName As String
    Dim sPivotFieldProperty As String
    Dim vPivotFieldValue As Variant
    
    Set pvt = This.sht.PivotTables(1)

    'Read the Orientation field of the pivot Fields into StorageRecords array
    m001_DataAccess.PR_GetPivotFieldDataSubtotalProperty This.sht.Name, This.StorageRecords
    
    For i = LBound(This.StorageRecords) To UBound(This.StorageRecords)
        sPivotFieldName = This.StorageRecords(i).Name
        sPivotFieldProperty = This.StorageRecords(i).Property
        vPivotFieldValue = This.StorageRecords(i).Value
        pvt.PivotFields(sPivotFieldName).Subtotals(1) = vPivotFieldValue
    Next i

End Sub


Private Sub PR_SetPivotFieldProperties()

    Dim pvt As PivotTable
    Dim i As Long
    Dim sPivotFieldName As String
    Dim sPivotFieldProperty As String
    Dim vPivotFieldValue As Variant
    
    Set pvt = This.sht.PivotTables(1)

    'Read the Orientation field of the pivot Fields into StorageRecords array
    m001_DataAccess.PR_GetPivotFieldDataPropertiesExSubtotals This.sht.Name, This.StorageRecords
    
    For i = LBound(This.StorageRecords) To UBound(This.StorageRecords)
        sPivotFieldName = This.StorageRecords(i).Name
        sPivotFieldProperty = This.StorageRecords(i).Property
        vPivotFieldValue = This.StorageRecords(i).Value
        CallByName pvt.PivotFields(sPivotFieldName), sPivotFieldProperty, VbLet, vPivotFieldValue
    Next i

End Sub

Sub PR_SetPivotTableLocation()
'Move pivot table down now that filters are in place
    
    This.sht.Range("1:5").Insert xlShiftDown

End Sub


Private Sub PR_SetSheetFormat()
'Sets sht format includiing range names
    
    With This.sht
        
        'Necessary to activate to set gridlines and zoom
        .Activate
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.Zoom = 80

        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 11
        .DisplayPageBreaks = False
        .Columns("A:A").ColumnWidth = 4
        
        'Add sheet heading named range and format
        .Names.Add Name:="SheetHeading", RefersTo:="=$B$2"
        .Range("SheetHeading") = m001_DataAccess.PR_GetHeadingNameBasedOnSheetName(This.sht.Name)
        .Range("SheetHeading").Font.Bold = True
        .Range("SheetHeading").Font.Size = 16
        
        'Add sheet category named range and format
        .Names.Add Name:="SheetCategory", RefersTo:="=$A$1"
        .Range("SheetCategory") = m001_DataAccess.PR_GetCategoryNameBasedOnSheetName(This.sht.Name)
        .Range("SheetCategory").Font.Color = RGB(170, 170, 170)
        .Range("SheetCategory").Font.Size = 8

    End With

End Sub



Private Sub PR_SetViewLayout()
'Freezes panes and sets the column widths.  Can only be set once the rest of the pivot table is
'in place
    
    Dim vFreezeDetails As Variant
    Dim iFreezeRow As Integer
    Dim iFreezeCol As Integer
    Dim vRowRangeColWidths As Variant
    Dim vArrayRowRangeColWidths As Variant
    Dim vDataBodyRangeColWidth As Variant
    Dim i As Integer
    
    'Set freeze pane locations
    vFreezeDetails = m001_DataAccess.PR_GetFreezePaneLocation(This.sht.Name)
    If Not IsNull(vFreezeDetails) Then
        iFreezeRow = Split(vFreezeDetails, "|")(0)
        iFreezeCol = Split(vFreezeDetails, "|")(1)
        This.sht.Activate
        This.sht.Cells(iFreezeRow + 1, iFreezeCol + 1).Select
        ActiveWindow.FreezePanes = True
    End If

    'Set Row Range column widths
    vRowRangeColWidths = m001_DataAccess.PR_GetRowRangeColWidths(This.sht.Name)
    If Not IsNull(vRowRangeColWidths) Then
        vArrayRowRangeColWidths = Split(vRowRangeColWidths, "|")
        For i = LBound(vArrayRowRangeColWidths) To UBound(vArrayRowRangeColWidths)
            This.sht.PivotTables(1).RowRange.Rows(1).Cells(i + 1).ColumnWidth = _
                vArrayRowRangeColWidths(i)
        Next i
    End If

    'Set data body range column widths (all widths are set identical sizes)
    vDataBodyRangeColWidth = m001_DataAccess.PR_GetDataBodyRowRangeColWidth(This.sht.Name)
    If Not IsNull(vDataBodyRangeColWidth) Then
        This.sht.PivotTables(1).DataBodyRange.EntireColumn.ColumnWidth = vDataBodyRangeColWidth
    End If


End Sub



Private Function PR_PivotTableHasRowRange() As Boolean

    Dim rng As Range
    On Error Resume Next
    
    Set rng = This.sht.PivotTables(1).RowRange
    
    If Err.Number = 0 Then
        PR_PivotTableHasRowRange = True
    Else
        PR_PivotTableHasRowRange = False
    End If

    On Error GoTo 0

End Function


Private Function ReportBackingQuery() As Variant
'Returns a DAX query for the current power query report to provide backing data.  Null is returned if
'it doesnt exist
    
    
    
    
    

End Function



Private Function PR_PivotTableHasDataBodyRange() As Boolean

    Dim rng As Range
    On Error Resume Next
    
    Set rng = This.sht.PivotTables(1).DataBodyRange
    
    If Err.Number = 0 Then
        PR_PivotTableHasDataBodyRange = True
    Else
        PR_PivotTableHasDataBodyRange = False
    End If

    On Error GoTo 0

End Function


Private Sub CreateDummyPowerPivotTable(ByVal sQueryName As String)
'Creates empty table loaded into Power Pivot to form source of listobject connection for DAX queries

    Dim sQueryStr As String
    Dim sTestTableName As String
    Dim bDummyTableExistsInDataModel As Boolean
    
    
    On Error Resume Next
    sTestTableName = ActiveWorkbook.Model.ModelTables(sQueryName).Name
    bDummyTableExistsInDataModel = (Err.Number = 0)
    On Error GoTo 0
    
    If Not bDummyTableExistsInDataModel Then
    
        'Query may exist but not in data model
        On Error Resume Next
            ActiveWorkbook.Queries(sQueryName).Delete
        On Error GoTo 0
    
        sQueryStr = "let" & vbCr & _
            "    tbl = Table.FromRecords({[NullHeader = """"]}), " & vbCr & _
            "    ChangedType = Table.TransformColumnTypes(tbl, {{""NullHeader"", type text}})" & vbCr & _
            "in" & vbCr & _
            "    ChangedType"
    
        ActiveWorkbook.Queries.Add Name:="DummyTableForDaxQueries", Formula:=sQueryStr
        Workbooks("PivotReportExample.xlsm").Connections.Add2 _
            Name:="Query - DummyTableForDaxQueries", _
            Description:="Connection to the 'DummyTableForDaxQueries' query in the workbook.", _
            ConnectionString:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=DummyTableForDaxQueries;Extended Properties=", _
            CommandText:="""DummyTableForDaxQueries""", _
            lCmdtype:=6, _
            CreateModelConnection:=True, _
            ImportRelationships:=False

    End If

End Sub



Private Function PR_CreateDataQueryBackingSheet() As Worksheet
'Sets sht format includiing range names
    
    Dim shtDataBacking As Worksheet
    Dim wkb As Workbook
    
    
    Set wkb = This.sht.Parent
    On Error Resume Next
    wkb.Sheets(This.sht.Name & "_Backing").Delete
    On Error GoTo 0
    
    Set shtDataBacking = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    shtDataBacking.Name = This.sht.Name & "_Backing"
    
    With shtDataBacking
        
        'Necessary to activate to set gridlines and zoom
        .Activate
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.Zoom = 80

        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 11
        .DisplayPageBreaks = False
        .Columns("A:A").ColumnWidth = 4
        
        'Add sheet heading named range and format
        .Names.Add Name:="SheetHeading", RefersTo:="=$B$2"
        .Range("SheetHeading") = "Backing: " & _
            m001_DataAccess.PR_GetHeadingNameBasedOnSheetName(This.sht.Name)
        .Range("SheetHeading").Font.Bold = True
        .Range("SheetHeading").Font.Size = 16
        
        'Add sheet category named range and format
        .Names.Add Name:="SheetCategory", RefersTo:="=$A$1"
        .Range("SheetCategory") = m001_DataAccess.PR_GetCategoryNameBasedOnSheetName(This.sht.Name)
        .Range("SheetCategory").Font.Color = RGB(170, 170, 170)
        .Range("SheetCategory").Font.Size = 8

    End With

    Set PR_CreateDataQueryBackingSheet = shtDataBacking

End Function


Private Sub PR_CreateListObjectBacking(ByVal sht As Worksheet, _
    ByVal sQueryName As String, ByVal sReportName As String, _
    ByVal sDaxQuery As String)
'Utilsied for creating optional backing tab of data when values report sheet is created.
'This sub creates the ListObject query.


    Dim lo As ListObject
    Dim sConnectionName As String
    
    'PowerPivot seems to append "Query - " to the query name to arrive at connection name
    sConnectionName = "Query - " & sQueryName
    
    Set lo = ActiveSheet.ListObjects.Add( _
        SourceType:=xlSrcModel, _
        Source:=ActiveWorkbook.Connections(sConnectionName), _
        Destination:=Range("$B$5"))
             
    With lo.TableObject
        .PreserveFormatting = True
        .RefreshStyle = xlOverwriteCells
        .AdjustColumnWidth = True
        .ListObject.Name = "tbl_Backing"
    End With
    
    With lo.TableObject.WorkbookConnection.OLEDBConnection
        .CommandText = Array(sDaxQuery)
        .CommandType = xlCmdDAX
    End With

    lo.Refresh
    
    'Remove link to Power Pivot change to values
    lo.TableObject.Delete
    
    
End Sub

