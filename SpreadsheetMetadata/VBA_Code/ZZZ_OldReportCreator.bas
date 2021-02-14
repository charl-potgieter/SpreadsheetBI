Attribute VB_Name = "ZZZ_OldReportCreator"
'Option Explicit
'Option Private Module
'
'
'
'Sub CreatePivotTable(ByVal sSheetName As String, ByRef pvt As PivotTable)
'
'    Dim loPivotTableSettings As ListObject
'    Dim loPivotFieldSettings As ListObject
'    Dim sht As Worksheet
'
'    If SheetExists(ActiveWorkbook, sSheetName) Then
'        MsgBox ("deleting sheet, need to give user a choice")
'        ActiveWorkbook.Sheets(sSheetName).Delete
'    End If
'
'    Set sht = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
'    sht.Name = sSheetName
'
'    'Create pivot in first row and then shift down.  This is easiest approach to get correct location.  Rows are inserted in calling sub
'    'once pivot design is complete
'
'    Set pvt = ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
'        ActiveWorkbook.Connections("ThisWorkbookDataModel"), Version:=6). _
'        CreatePivotTable(sht.Range("B1"))
'
'End Sub
'
'
'Sub CustomisePivotTable(ByRef pvt As PivotTable, ReportProperties As TypeReportProperties)
'
'    With ReportProperties
'        pvt.HasAutoFormat = .AutoFit
'        pvt.ColumnGrand = .ColumnTotals
'        pvt.RowGrand = .RowTotals
'        pvt.ShowDrillIndicators = .DisplayExpandButtons
'        pvt.DisplayFieldCaptions = .DisplayFieldHeaders
'    End With
'
'End Sub
'
'Sub SetPivotFields(ByRef pvt As PivotTable, ByRef ReportFieldSettings() As TypeReportFieldSettings)
'
'    Dim i As Integer
'    Dim j As Integer
'
'    For i = 0 To UBound(ReportFieldSettings)
'        With ReportFieldSettings(i)
'
'            'Set field orientation
'            Select Case True
'                Case .FieldType = "Measure"
'                    pvt.CubeFields(.CubeFieldName).Orientation = xlDataField
'                Case .Orientation = "Row"
'                    pvt.CubeFields(.CubeFieldName).Orientation = xlRowField
'                Case .Orientation = "Column"
'                    pvt.CubeFields(.CubeFieldName).Orientation = xlColumnField
'                Case .Orientation = "Filter"
'                    pvt.CubeFields(.CubeFieldName).Orientation = xlPageField
'            End Select
'
'            'Filter columns, rows, filters if required
'            Select Case True
'                Case .FieldType = "Measure" Or Not (ArrayIsDimensioned(.FilterValues))
'                    'Do Nothing
'                Case .FilterType = "Include"
'                    PivotFilterInclude pvt, .CubeFieldName, .FilterValues
'                Case .FilterType = "Exclude"
'                    PivotFilterExclude pvt, .CubeFieldName, .FilterValues
'            End Select
'
'            'Set subtotals
'            Select Case True
'                Case .FieldType = "Measure"
'                    'Do Nothing
'                Case .Subtotal
'                    ShowPivotMeasureSubtotals pvt, .CubeFieldName
'                Case Not (.Subtotal)
'                    HidePivotMeasureSubtotals pvt, .CubeFieldName
'            End Select
'
'            'Set subtotal location
'            Select Case True
'                Case .FieldType = "Measure" Or Not (.Subtotal)
'                    'Do Nothing
'                Case .SubtotalAtTop
'                    SetPivotSubtotalsAtTop pvt, .CubeFieldName
'                Case Not (.SubtotalAtTop)
'                    SetPivotSubtotalsAtBottom pvt, .CubeFieldName
'            End Select
'
'            'Insert blank line if appropriate
'            If .FieldType = "Column" And .BlankLine Then
'                pvt.CubeFields(.CubeFieldName).PivotFields(1).LayoutBlankLine = True
'            End If
'
'            'Format measures
'            Select Case True
'                Case .FieldType <> "Measure"
'                    'Do Nothing
'                Case .CustomFormat <> ""
'                    pvt.PivotFields(.CubeFieldName).NumberFormat = .CustomFormat
'                Case .Format = "Zero Decimals"
'                    pvt.PivotFields(.CubeFieldName).NumberFormat = "#,##0_);(#,##0);-??"
'                Case .Format = "One Decimal"
'                    pvt.PivotFields(.CubeFieldName).NumberFormat = "#,##0.0_);(#,##0.0);-??"
'                Case .Format = "Two Decimals"
'                    pvt.PivotFields(.CubeFieldName).NumberFormat = "#,##0.00_);(#,##0.00);-??"
'            End Select
'
'
'        End With
'    Next i
'
'
'    'Need to handle collapsing fields on a seperate loop as cannot collapse before underlying rows or
'    'columns are in place
'    For i = 0 To UBound(ReportFieldSettings)
'        With ReportFieldSettings(i)
'            If (.FieldType <> "Measure") And (ArrayIsDimensioned(.CollapseFieldValues)) And (.Orientation <> "Filter") Then
'                CollapsePivotFieldValues pvt, .CubeFieldName, .CollapseFieldValues
'            End If
'        End With
'    Next i
'
'
'
'
'
'
'
'
'
'End Sub
'
'
'
'
'
'Sub PivotFilterInclude(ByRef pvt As PivotTable, ByVal sCubeFieldName, ByRef aIncludedItems() As String)
''Filter sCubeFieldNamr in pvt to for items in aIncludedItems
'
'    Dim pf As PivotField
'    Dim sPivotFieldName As String
'    Dim aArrayVariantOfIncludedItems() As Variant
'    Dim i As Integer
'
'    If pvt.CubeFields(sCubeFieldName).Orientation = xlPageField Then
'        pvt.CubeFields(sCubeFieldName).EnableMultiplePageItems = True
'    End If
'
'    sPivotFieldName = pvt.CubeFields(sCubeFieldName).PivotFields(1).Name
'    Set pf = pvt.PivotFields(sPivotFieldName)
'
'    ReDim aArrayVariantOfIncludedItems(UBound(aIncludedItems))
'    For i = LBound(aIncludedItems) To UBound(aIncludedItems)
'        aArrayVariantOfIncludedItems(i) = sCubeFieldName & ".&[" & aIncludedItems(i) & "]"
'    Next i
'
'    pf.CubeField.IncludeNewItemsInFilter = False
'    pf.VisibleItemsList = Array(aArrayVariantOfIncludedItems)
'
'End Sub
'
'
'
'Sub PivotFilterExclude(ByRef pvt As PivotTable, ByVal sCubeFieldName, ByRef aExcludedItems() As String)
''Filter sCubeFieldNamr in pvt to to exclude items in aExcludedItems
'
'    Dim pf As PivotField
'    Dim sPivotFieldName As String
'    Dim aArrayVariantOfExcludedItems() As Variant
'    Dim aTempArrayIncludedItems() As String
'    Dim item As Variant
'    Dim sItemExclField As String
'    Dim i As Integer
'
'
'    sPivotFieldName = pvt.CubeFields(sCubeFieldName).PivotFields(1).Name
'    Set pf = pvt.PivotFields(sPivotFieldName)
'
'    If pf.Orientation <> xlPageField Then
'        ReDim aArrayVariantOfExcludedItems(UBound(aExcludedItems))
'        For i = LBound(aExcludedItems) To UBound(aExcludedItems)
'            aArrayVariantOfExcludedItems(i) = sCubeFieldName & ".&[" & aExcludedItems(i) & "]"
'        Next i
'        pf.CubeField.IncludeNewItemsInFilter = True
'        pf.HiddenItemsList = Array(aArrayVariantOfExcludedItems)
'    Else
'        'Cannot appply the above technique on page fields for cubefields for some reason.
'        'below is best workaround I can come up with (temporary switch to row field and set
'        'visible items to exclude items to be filtered out, then switch back to page field)
'        pvt.CubeFields(sCubeFieldName).EnableMultiplePageItems = True
'        pvt.CubeFields(sCubeFieldName).Orientation = xlRowField
'        i = 0
'        For Each item In pf.VisibleItems
'            If i = 0 Then
'                ReDim aTempArrayIncludedItems(i)
'            Else
'                ReDim Preserve aTempArrayIncludedItems(i)
'            End If
'            'Add items not in excluded items into the Included Items array
'            sItemExclField = Replace(Replace(Replace(item, sCubeFieldName & ".&", ""), "[", ""), "]", "")
'            If Not (ValueIsInStringArray(sItemExclField, aExcludedItems)) Then
'                aTempArrayIncludedItems(i) = item
'                i = i + 1
'            End If
'        Next item
'        pvt.CubeFields(sCubeFieldName).Orientation = xlPageField
'        pf.CubeField.IncludeNewItemsInFilter = False
'        pf.VisibleItemsList = aTempArrayIncludedItems
'    End If
'
'End Sub
'
'
'Sub ShowPivotMeasureSubtotals(ByRef pvt As PivotTable, ByVal sCubeFieldName As String)
'
'    Dim pf As PivotField
'
'    Set pf = pvt.CubeFields(sCubeFieldName).PivotFields(1)
'    pf.Subtotals = Array(True, False, False, False, False, False, False, _
'        False, False, False, False, False)
'
'End Sub
'
'
'Sub HidePivotMeasureSubtotals(ByRef pvt As PivotTable, ByVal sCubeFieldName As String)
'
'    Dim pf As PivotField
'
'    Set pf = pvt.CubeFields(sCubeFieldName).PivotFields(1)
'    pf.Subtotals = Array(False, False, False, False, False, False, False, _
'        False, False, False, False, False)
'
'End Sub
'
'
'Sub SetPivotSubtotalsAtTop(ByRef pvt As PivotTable, ByVal sCubeFieldName As String)
'
'    Dim pf As PivotField
'
'    Set pf = pvt.CubeFields(sCubeFieldName).PivotFields(1)
'    pf.LayoutSubtotalLocation = xlAtTop
'
'End Sub
'
'
'Sub SetPivotSubtotalsAtBottom(ByRef pvt As PivotTable, ByVal sCubeFieldName As String)
'
'    Dim pf As PivotField
'
'    Set pf = pvt.CubeFields(sCubeFieldName).PivotFields(1)
'    pf.LayoutSubtotalLocation = xlAtBottom
'
'End Sub
'
'
'
'Sub CollapsePivotFieldValues(ByRef pvt As PivotTable, ByVal sCubeFieldName, ByRef aCollapseValues() As String)
''Collapse rows / columns of sCubeFieldName where they have values equal to those in aCollapseValues array
'
'    Dim pf As PivotField
'    Dim pi As PivotItem
'    Dim sPivotItemName As String
'    Dim sItemExclField As String
'    Dim i As Integer
'
'    Set pf = pvt.CubeFields(sCubeFieldName).PivotFields(1)
'
'    'Cannot simply loop all values to collapse and if multiple need to be collapse, collapsing the
'    '2nd item makes the first visible.  Solution is to hide all and then selectively unhide if more
'    'than one item needs to be collapsed
'    If UBound(aCollapseValues) = 0 Then
'        Set pf = pvt.CubeFields(sCubeFieldName).PivotFields(1)
'        sPivotItemName = sCubeFieldName & ".&[" & aCollapseValues(0) & "]"
'        Set pi = pf.PivotItems(sPivotItemName)
'        pi.DrilledDown = False
'    Else
'        pf.DrilledDown = False
'        For Each pi In pf.PivotItems
'            sItemExclField = Replace(Replace(Replace(pi.Value, sCubeFieldName & ".&", ""), "[", ""), "]", "")
'            If Not ValueIsInStringArray(sItemExclField, aCollapseValues) Then
'                pi.DrilledDown = True
'            End If
'        Next pi
'    End If
'
'
'
'End Sub
'
'
