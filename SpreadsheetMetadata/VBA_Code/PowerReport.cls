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

'Utilised for saving data
Private Type TypePowerReportStorageRecord
    SheetName As String
    Name As String
    DataType As String
    Property As String
    Value As String
End Type

'Use "This" declaration as an easy way to get intellisense to the classes private variables
'https://rubberduckvba.wordpress.com/2020/02/27/vba-classes-gateway-to-solid/
Private Type TypePowerReport
    SheetName As String
    SheetHeading As String
    SheetCategory As String
    PvtTable As PivotTable
    StorageRecords() As TypePowerReportStorageRecord
    StorageIndex As Long
End Type
    
Private This As TypePowerReport
Private Const MaxStorageRecords As Long = 1000000


Public Sub AssignToExistingSheet(ByRef sht As Worksheet)
    This.SheetName = sht.Name
    This.SheetHeading = sht.Range("SheetHeading")
    This.SheetCategory = sht.Range("SheetCategory")
    Set This.PvtTable = sht.PivotTables(1)
End Sub




Public Sub CreateFromStoredData(ByRef wkb As Workbook, ByVal sSheetName)
'Note this will delete any existing sheet with same name
'sSheetName needs to already exist in stored data
    
    Dim sht As Worksheet
    Dim pvt As PivotTable
    
    On Error Resume Next
    wkb.Sheets(sSheetName).Delete
    On Error GoTo 0

    'Setup the Worksheet
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    sht.Name = sSheetName
    FormatReportSheet sht
    sht.Range("SheetHeading") = m001_DataAccess.GetReportHeadingNameBasedOnSheetName(sSheetName)
    sht.Range("SheetCategory") = m001_DataAccess.GetReportCategoryNameBasedOnSheetName(sSheetName)
    
    'Create empty pivot table from data model
    Set pvt = ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
        ActiveWorkbook.Connections("ThisWorkbookDataModel"), Version:=6). _
        CreatePivotTable(sht.Range("B7"))
    

End Sub

Public Sub SavePowerReportStructure()
    
    m001_DataAccess.SetupOrAssignPowerReportStorage
    m001_DataAccess.DeleteExistingReportData This.SheetName
    ReDim This.StorageRecords(MaxStorageRecords)
    
    This.StorageIndex = 0
    ReadPowerReportSheetProperties
    ReadPowerReportPivotTableProperties
    ReadPowerReportCubeFieldProperties
    ReadPowerReportPivotFieldProperties
    ReDim Preserve This.StorageRecords(This.StorageIndex)
    WritePowerReportPropertiesToStorage
    
End Sub



Private Sub ReadPowerReportSheetProperties()
'Reads sheet properties into records for later saving

    'Read Sheet Heading
    With This.StorageRecords(This.StorageIndex)
        .SheetName = This.SheetName
        .Name = This.SheetName
        .DataType = "SheetDataType"
        .Property = "SheetHeading"
        .Value = This.SheetHeading
    End With
    
    'Read Sheet Category
    This.StorageIndex = This.StorageIndex + 1
    With This.StorageRecords(This.StorageIndex)
        .SheetName = This.SheetName
        .Name = This.SheetName
        .DataType = "SheetDataType"
        .Property = "SheetCategory"
        .Value = This.SheetCategory
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
                .SheetName = This.SheetName
                .Name = This.PvtTable.Name
                .DataType = "PivotTableDataType"
                .Property = item
                .Value = CallByName(This.PvtTable, item, VbGet)
            End With
        Next item
    
End Sub

Private Sub ReadPowerReportCubeFieldProperties()
'Some of Object model properties are set at cube field level.   Others at a PivotField level
 
    Dim pvtCubeField As cubeField
    Dim Properties As Variant
    Dim item As Variant
    Dim i As Long

    Properties = Array("Name", "Caption", "Orientation", "Position")

    For Each pvtCubeField In This.PvtTable.CubeFields
        If pvtCubeField.Orientation <> xlHidden Then
            For Each item In Properties
                This.StorageIndex = This.StorageIndex + 1
                With This.StorageRecords(This.StorageIndex)
                    .SheetName = This.SheetName
                    .Name = pvtCubeField.Name
                    .DataType = "PivotCubeFieldDataType"
                    .Property = item
                    .Value = CallByName(pvtCubeField, item, VbGet)
                End With
            Next item
        End If
    Next pvtCubeField


End Sub


Private Sub ReadPowerReportPivotFieldProperties()
'Some of Object model properties are set at pivot field level.   Others at a cubefield level
'Simpler not to loop a Properties array due to exceptions that need to be handled as below.

    Dim pvtField As PivotField
    Dim i As Long

    For Each pvtField In This.PvtTable.PivotFields

        If pvtField.Name = "Values" Then
            'Special case.  This is only case where orientation is set at a pivot
            'field level.  Other cases are set at cube field level
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.SheetName
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "Orientation"
                .Value = pvtField.Orientation
            End With

        ElseIf pvtField.Orientation = xlDataField Then
                
            'SubtotalName
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.SheetName
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "SubtotalName"
                .Value = pvtField.SubtotalName
            End With
            
            'Subotal(note the property is indexed)
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.SheetName
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "Subtotals"
                'Note the index one here to designate automatic for powerpivot
                .Value = pvtField.Subtotals(1)
            End With
            
            'LayoutSubtotalLocation
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.SheetName
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "LayoutSubtotalLocation"
                .Value = pvtField.LayoutSubtotalLocation
            End With
            
            'NumberFormat
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.SheetName
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "NumberFormat"
                .Value = pvtField.NumberFormat
            End With

        Else
        
            'LayoutBlankLine
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.SheetName
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "LayoutBlankLine"
                .Value = pvtField.LayoutBlankLine
            End With
            
            'LayoutCompactRow
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.SheetName
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "LayoutCompactRow"
                .Value = pvtField.LayoutCompactRow
            End With
            
            'LayoutForm
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.SheetName
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "LayoutForm"
                .Value = pvtField.LayoutForm
            End With
            
            'LayoutPageBreak
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.SheetName
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "LayoutPageBreak"
                .Value = pvtField.LayoutPageBreak
            End With
            
            'RepeatLables
            This.StorageIndex = This.StorageIndex + 1
            With This.StorageRecords(This.StorageIndex)
                .SheetName = This.SheetName
                .Name = pvtField.Name
                .DataType = "PivotFieldDataType"
                .Property = "RepeatLabels"
                .Value = pvtField.RepeatLabels
            End With
            
        End If

    Next pvtField

End Sub

'********************************************************************************************************
'
'                               Private Subs
'
'********************************************************************************************************

Private Sub WritePowerReportPropertiesToStorage()

    Dim i As Long

    For i = LBound(This.StorageRecords) To UBound(This.StorageRecords)
        m001_DataAccess.WritePowerReportRecord _
            This.StorageRecords(i).SheetName, This.StorageRecords(i).Name, _
            This.StorageRecords(i).DataType, This.StorageRecords(i).Property, _
            This.StorageRecords(i).Value
    Next i

End Sub



Private Sub FormatReportSheet(ByRef sht As Worksheet)
'Applies my preferred sheet formattting

    sht.Activate
    
    sht.Cells.Font.Name = "Calibri"
    sht.Cells.Font.Size = 11
    
    sht.Range("A1").Font.Color = RGB(170, 170, 170)
    sht.Range("A1").Font.Size = 8
    
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 80
    sht.DisplayPageBreaks = False
    sht.Columns("A:A").ColumnWidth = 4
    
    If SheetLevelRangeNameExists(sht, "SheetHeading") Then
        sht.Names("SheetHeading").Delete
    End If
    sht.Names.Add Name:="SheetHeading", RefersTo:="=$B$2"
    
    If SheetLevelRangeNameExists(sht, "SheetCategory") Then
        sht.Names("SheetCategory").Delete
    End If
    sht.Names.Add Name:="SheetCategory", RefersTo:="=$A$1"
    
    With sht.Range("SheetHeading")
        If .Value = "" Then
            .Value = "Heading"
        End If
        .Font.Bold = True
        .Font.Size = 16
    End With

End Sub



