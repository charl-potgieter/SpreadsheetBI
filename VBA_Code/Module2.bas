Attribute VB_Name = "Module2"
Option Explicit

Sub Macro1()
'
' Macro1 Macro
'

'
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
        ActiveWorkbook.Connections("ThisWorkbookDataModel"), Version:=6). _
        CreatePivotTable TableDestination:="Sheet2!R6C2", TableName:="PivotTable1" _
        , DefaultVersion:=6
    Cells(6, 2).Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = True
        .CompactRowIndent = 1
        .VisualTotals = False
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = True
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .DisplayEmptyRow = False
        .DisplayEmptyColumn = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .DisplayImmediateItems = True
        .ViewCalculatedMembers = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = True
        .RowAxisLayout xlCompactRow
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotCache.RefreshOnFileOpen = False
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable1").CubeFields( _
        "[tbl_LatestInstance].[Description]")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").CubeFields("[Measures].[MyCode]")
    With ActiveSheet.PivotTables("PivotTable1").CubeFields( _
        "[tbl_LatestInstance].[Date]")
        .Orientation = xlColumnField
        .Position = 1
    End With
End Sub
