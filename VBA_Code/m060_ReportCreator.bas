Attribute VB_Name = "m060_ReportCreator"
Option Explicit
Option Private Module



Sub CreatePivotTable(ByVal sSheetName As String, ByRef pvt As PivotTable)

    Dim loPivotTableSettings As ListObject
    Dim loPivotFieldSettings As ListObject
    Dim sht As Worksheet

    If SheetExists(ActiveWorkbook, sSheetName) Then
        MsgBox ("deleting sheet, need to give user a choice")
        ActiveWorkbook.Sheets(sSheetName).Delete
    End If

    Set sht = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    sht.Name = sSheetName
    
    'Create pivot in first row and then shift down.  This is easiest approach to get correct location.  Rows are inserted in calling sub
    'once pivot design is complete
    Set pvt = ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
        ActiveWorkbook.Connections("ThisWorkbookDataModel"), Version:=6). _
        CreatePivotTable(sht.Range("B1"))
    


    
    
End Sub


Sub CustomisePivotTable(ByRef pvt As PivotTable, ReportProperties As TypeReportProperties)

    
    pvt.HasAutoFormat = ReportProperties.AutoFit
    pvt.ColumnGrand = ReportProperties.ColumnTotals
    pvt.RowGrand = ReportProperties.RowTotals
    

End Sub

Sub SetPivotFields(ByRef pvt As PivotTable, ByRef ReportFieldSettings() As TypeReportFieldSettings)

    Dim i As Integer
    Dim j As Integer

    
    For i = 0 To UBound(ReportFieldSettings)
    
        
        With ReportFieldSettings(i)
            
            'Set field orientation
            Select Case True
                Case .FieldType = "Measure"
                    pvt.CubeFields(.CubeFieldName).Orientation = xlDataField
                Case .Orientation = "Row"
                    pvt.CubeFields(.CubeFieldName).Orientation = xlRowField
                Case .Orientation = "Column"
                    pvt.CubeFields(.CubeFieldName).Orientation = xlColumnField
                Case .Orientation = "Filter"
                    pvt.CubeFields(.CubeFieldName).Orientation = xlPageField
            End Select
            
            'Filter columns, rows, filters if required
            Select Case .FilterType
                Case "Include"
                    PivotFilterInclude pvt, .CubeFieldName, .FilterValues
                Case "Exclude"
                    PivotFilterExclude pvt, .CubeFieldName, .FilterValues
            End Select
            
        End With
          
        
    Next i
    
    
                
                
    'Set field on either data, row of colum
    
    
    'Format field
'    Select Case .ListColumns("Format").DataBodyRange.Cells(i)
'        Case "Zero Decimals"
'            pvt.PivotFields(sCubeFieldName).NumberFormat = "#,##0_);(#,##0);-??"
'        Case "One Decimal"
'            pvt.PivotFields(sCubeFieldName).NumberFormat = "#,##0.0_);(#,##0.0);-??"
'        Case "Two Decimals"
'            pvt.PivotFields(sCubeFieldName).NumberFormat = "#,##0.00_);(#,##0.00);-??"
'    End Select
                

End Sub





Sub PivotFilterInclude(ByRef pvt As PivotTable, ByVal sCubeFieldName, ByRef aIncludedItems() As String)
'Filter sCubeFieldNamr in pvt to for items in aIncludedItems

    Dim pf As PivotField
    Dim sPivotFieldName As String
    Dim aArrayVariantOfIncludedItems() As Variant
    Dim i As Integer

    If pvt.CubeFields(sCubeFieldName).Orientation = xlPageField Then
        pvt.CubeFields(sCubeFieldName).EnableMultiplePageItems = True
    End If
    
    sPivotFieldName = pvt.CubeFields(sCubeFieldName).PivotFields(1).Name
    Set pf = pvt.PivotFields(sPivotFieldName)
    
    ReDim aArrayVariantOfIncludedItems(UBound(aIncludedItems))
    For i = LBound(aIncludedItems) To UBound(aIncludedItems)
        aArrayVariantOfIncludedItems(i) = sCubeFieldName & ".&[" & aIncludedItems(i) & "]"
    Next i

    pf.CubeField.IncludeNewItemsInFilter = False
    pf.VisibleItemsList = Array(aArrayVariantOfIncludedItems)

End Sub



Sub PivotFilterExclude(ByRef pvt As PivotTable, ByVal sCubeFieldName, ByRef aExcludedItems() As String)
'Filter sCubeFieldNamr in pvt to to exclude items in aExcludedItems

    Dim pf As PivotField
    Dim sPivotFieldName As String
    Dim aArrayVariantOfExcludedItems() As Variant
    Dim aTempArrayIncludedItems() As String
    Dim item As Variant
    Dim sItemExclField As String
    Dim i As Integer

    
    sPivotFieldName = pvt.CubeFields(sCubeFieldName).PivotFields(1).Name
    Set pf = pvt.PivotFields(sPivotFieldName)
    
    If pf.Orientation <> xlPageField Then
        ReDim aArrayVariantOfExcludedItems(UBound(aExcludedItems))
        For i = LBound(aExcludedItems) To UBound(aExcludedItems)
            aArrayVariantOfExcludedItems(i) = sCubeFieldName & ".&[" & aExcludedItems(i) & "]"
        Next i
        pf.CubeField.IncludeNewItemsInFilter = True
        pf.HiddenItemsList = Array(aArrayVariantOfExcludedItems)
    Else
        'Cannot appply the above technique on page fields for cubefields for some reason.
        'below is best workaround I can come up with (temporary switch to row field and set
        'visible items to exclude items to be filtered out, then switch back to page field)
        pvt.CubeFields(sCubeFieldName).EnableMultiplePageItems = True
        pvt.CubeFields(sCubeFieldName).Orientation = xlRowField
        i = 0
        For Each item In pf.VisibleItems
            If i = 0 Then
                ReDim aTempArrayIncludedItems(i)
            Else
                ReDim Preserve aTempArrayIncludedItems(i)
            End If
            'Add items not in excluded items into the Included Items array
            sItemExclField = Replace(Replace(Replace(item, sCubeFieldName & ".&", ""), "[", ""), "]", "")
            If Not (ValueIsInStringArray(sItemExclField, aExcludedItems)) Then
                aTempArrayIncludedItems(i) = item
                i = i + 1
            End If
        Next item
        pvt.CubeFields(sCubeFieldName).Orientation = xlPageField
        pf.CubeField.IncludeNewItemsInFilter = False
        pf.VisibleItemsList = aTempArrayIncludedItems
    End If
        
End Sub



