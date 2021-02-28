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


Public Sub AssignToExistingSheet(ByRef sht As Worksheet)
    Set This.sht = sht
End Sub


Public Sub SavePowerReportStructure()
    
    m001_DataAccess.PR_SetupOrAssignStorage
    m001_DataAccess.PR_DeleteExistingData This.sht.Name
    ReDim This.StorageRecords(cPR_MaxStorageRecords)
    
    This.StorageIndex = 0
    ReadPowerReportSheetProperties
    ReadPowerReportPivotTableProperties
    ReadPowerReportCubeFieldProperties
    ReadPowerReportPivotFieldProperties
    ReDim Preserve This.StorageRecords(This.StorageIndex)
    WritePowerReportPropertiesToStorage
    
End Sub

Public Sub CreateFromStoredData(ByRef wkb As Workbook, ByVal sSheetName)
'Note this will delete any existing sheet with same name
'sSheetName needs to already exist in stored data
    
    Dim pvt As PivotTable
    
    Set This.sht = PR_CreateWorksheet(wkb, sSheetName)
    PR_CreateEmptyPivot
    PR_SetPivotTableProperties
    PR_SetCubeFieldOrientation
    PR_SetCubeFieldPropertiesExOrientation
    PR_SetPivotFieldSubtotals
    PR_SetPivotFieldProperties
    


End Sub

'********************************************************************************************************
'
'                               Private Subs
'
'********************************************************************************************************


Private Sub ReadPowerReportSheetProperties()
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


Private Sub ReadPowerReportPivotTableProperties()
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

Private Sub ReadPowerReportCubeFieldProperties()
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


Private Sub ReadPowerReportPivotFieldProperties()
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

Private Sub WritePowerReportPropertiesToStorage()

    Dim i As Long

    For i = LBound(This.StorageRecords) To UBound(This.StorageRecords)
        m001_DataAccess.PR_WriteRecords _
            This.StorageRecords(i).SheetName, This.StorageRecords(i).Name, _
            This.StorageRecords(i).DataType, This.StorageRecords(i).Property, _
            This.StorageRecords(i).Value, This.StorageRecords(i).CubeFieldPosition
    Next i

End Sub


Private Function PR_CreateWorksheet(ByRef wkb As Workbook, ByVal sSheetName) As Worksheet
'Note this will delete any existing sSheetName if it exists

    Dim sht As Worksheet

    On Error Resume Next
    wkb.Sheets(sSheetName).Delete
    On Error GoTo 0

    
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    sht.Name = sSheetName
    sht.Cells.Font.Name = "Calibri"
    sht.Cells.Font.Size = 11
    
    'Add heading and category
    sht.Names.Add Name:="SheetCategory", RefersTo:="=$A$1"
    sht.Names.Add Name:="SheetHeading", RefersTo:="=$B$2"
    sht.Range("SheetHeading") = m001_DataAccess.PR_GetHeadingNameBasedOnSheetName(sSheetName)
    
    sht.Range("SheetCategory") = m001_DataAccess.PR_GetCategoryNameBasedOnSheetName(sSheetName)
        With sht.Range("SheetHeading")
        .Font.Bold = True
        .Font.Size = 16
    End With
    
    With sht.Range("SheetCategory")
        .Font.Color = RGB(170, 170, 170)
        .Font.Size = 8
    End With
    
    sht.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 80
    sht.DisplayPageBreaks = False
    sht.Columns("A:A").ColumnWidth = 4
    
    Set PR_CreateWorksheet = sht

End Function



Private Function PR_CreateEmptyPivot()

    Dim wkb As Workbook
    
    Set wkb = This.sht.Parent

    'Create empty pivot table from data model
    Set PR_CreateEmptyPivot = wkb.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
        ActiveWorkbook.Connections("ThisWorkbookDataModel"), Version:=6). _
        CreatePivotTable(This.sht.Range("B7"))

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

Sub PR_SetPivotFieldSubtotals()
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


Sub PR_SetPivotFieldProperties()

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
