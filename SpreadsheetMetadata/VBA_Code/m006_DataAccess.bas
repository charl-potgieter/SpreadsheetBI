Attribute VB_Name = "m006_DataAccess"
Option Explicit
Option Private Module





'Sub SavePivotReportMetaData(ByRef pvtReportMetaData() As TypePivotReport)
''Saves pivot report metadata on sheeet in active workbook
'
'    Dim i As Integer
'
'    If Not SheetExists(ActiveWorkbook, "ReportSheetProperties") Then
'        CreateReportPropertiesSheet ActiveWorkbook
'    End If
'
'    For i = LBound(pvtReportMetaData) To UBound(pvtReportMetaData)
'        DeleteAnyPreExistingData pvtReportMetaData(i)
'        WriteSheetProperties pvtReportMetaData(i)
'        WritePivotTableProperties pvtReportMetaData(i)
'        WritePivotCubeFieldProperties pvtReportMetaData(i)
'        WritePivotFieldProperties pvtReportMetaData(i)
'    Next i
'
'    FormatReportSheetProperties
'
'End Sub
'
'
'
'
'
'Sub CreateReportPropertiesSheet(ByRef wkb As Workbook)
'
'    Dim sHeaders As String
'    Dim shtPivotTableProperties As Worksheet
'    Dim lo As ListObject
'
'    Set shtPivotTableProperties = InsertFormattedSheetIntoActiveWorkbook
'    With shtPivotTableProperties
'        .Name = "ReportSheetProperties"
'        .Range("SheetCategory").Value = "Report metadata"
'        .Range("SheetHeading").Value = "Report Sheet Properties"
'
'        'Create table
'        .Range("B6") = "SheetName"
'        .Range("C6") = "Name"
'        .Range("D6") = "DataType"
'        .Range("E6") = "Property"
'        .Range("F6") = "Value"
'        Set lo = .ListObjects.Add(xlSrcRange, Range("$B$6").CurrentRegion, , xlYes)
'
'    End With
'    lo.Name = "tbl_ReportProperties"
'    FormatTable lo
'
'
'End Sub
'
'Sub FormatReportSheetProperties()
'
'    ActiveWorkbook.Sheets("ReportSheetProperties").Activate
'
'    With ActiveWindow
'        If .FreezePanes Then .FreezePanes = False
'        .SplitColumn = 0
'        .SplitRow = 6
'        .FreezePanes = True
'    End With
'
'    With ActiveWorkbook.Sheets("ReportSheetProperties")
'        .ListObjects("tbl_ReportProperties"). _
'            DataBodyRange.HorizontalAlignment = xlLeft
'        Cells.EntireColumn.AutoFit
'    End With
'
'
'End Sub
'
'
'
'Sub DeleteAnyPreExistingData(ByRef pvtReportMetaData As TypePivotReport)
''Delete any info if it exist for this sheet name
'
'    Dim lo As ListObject
'    Dim iSheetNameColumnNumber As Integer
'
'    Set lo = ActiveWorkbook.Worksheets("ReportSheetProperties").ListObjects("tbl_ReportProperties")
'    iSheetNameColumnNumber = WorksheetFunction.Match("SheetName", lo.HeaderRowRange, 0)
'
'    lo.Range.AutoFilter field:=iSheetNameColumnNumber, _
'        Criteria1:=pvtReportMetaData.SheetName
'    On Error Resume Next
'    lo.DataBodyRange.SpecialCells(xlCellTypeVisible).EntireRow.Delete (xlShiftUp)
'    On Error GoTo 0
'    lo.AutoFilter.ShowAllData
'
'
'End Sub
'
'Sub WriteSingleReportSheetPropertyRecord(ByVal sSheetName As String, _
'    ByVal sName As String, ByVal sDataType As String, _
'    ByVal sProperty As String, ByVal val As Variant)
'
'    Dim lo As ListObject
'    Dim i As Double
'
'    Set lo = ActiveWorkbook.Sheets("ReportSheetProperties").ListObjects("tbl_ReportProperties")
'    AddOneRowToListObject lo
'    i = lo.DataBodyRange.Rows.Count
'    lo.ListColumns("SheetName").DataBodyRange.Cells(i) = sSheetName
'    lo.ListColumns("Name").DataBodyRange.Cells(i) = sName
'    lo.ListColumns("DataType").DataBodyRange.Cells(i) = sDataType
'    lo.ListColumns("Property").DataBodyRange.Cells(i) = sProperty
'    lo.ListColumns("Value").DataBodyRange.Cells(i) = val
'
'End Sub
'
'
'
'Sub WriteSheetProperties(ByRef pvtReportMetaData As TypePivotReport)
'
'    With pvtReportMetaData.ReportingSheet
'        WriteSingleReportSheetPropertyRecord pvtReportMetaData.SheetName, _
'            .Name, "SheetDataType", "SheetHeading", .SheetHeading
'        WriteSingleReportSheetPropertyRecord pvtReportMetaData.SheetName, _
'            .Name, "SheetDataType", "SheetCategory", .SheetCategory
'
'    End With
'
'End Sub
'
'
'Sub WritePivotTableProperties(ByRef pvtReportMetaData As TypePivotReport)
'
'    Dim key As Variant
'    Dim lo As ListObject
'    Dim i As Double
'
'    Set lo = ActiveWorkbook.Sheets("ReportSheetProperties").ListObjects("tbl_ReportProperties")
'
'    With pvtReportMetaData.PvtTable
'        For Each key In .Properties.Keys
'            AddOneRowToListObject lo
'            i = lo.DataBodyRange.Rows.Count
'            lo.ListColumns("SheetName").DataBodyRange.Cells(i) = pvtReportMetaData.SheetName
'            lo.ListColumns("Name").DataBodyRange.Cells(i) = .Name
'            lo.ListColumns("DataType").DataBodyRange.Cells(i) = "PivotTableDataType"
'            lo.ListColumns("Property").DataBodyRange.Cells(i) = key
'            lo.ListColumns("Value").DataBodyRange.Cells(i) = .Properties.item(key)
'        Next key
'    End With
'
'End Sub
'
'Sub WritePivotCubeFieldProperties(ByRef pvtReportMetaData As TypePivotReport)
'
'    Dim key As Variant
'    Dim lo As ListObject
'    Dim i As Double
'
'    Set lo = ActiveWorkbook.Sheets("ReportSheetProperties").ListObjects("tbl_ReportProperties")
'
'    For i = LBound(pvtReportMetaData.PvtCubeFields) To UBound(pvtReportMetaData.PvtCubeFields)
'        With pvtReportMetaData.PvtCubeFields(i)
'
'            WriteSingleReportSheetPropertyRecord pvtReportMetaData.SheetName, _
'               .Name, "PivotCubeFieldDataType", "Caption", .Caption
'
'            WriteSingleReportSheetPropertyRecord pvtReportMetaData.SheetName, _
'               .Name, "PivotCubeFieldDataType", "Orientation", .Orientation
'
'            WriteSingleReportSheetPropertyRecord pvtReportMetaData.SheetName, _
'               .Name, "PivotCubeFieldDataType", "Position", .Position
'
'        End With
'    Next i
'
'End Sub
'
'
'
'Sub WritePivotFieldProperties(ByRef pvtReportMetaData As TypePivotReport)
'
'    Dim lo As ListObject
'    Dim i As Double
'
'    Set lo = ActiveWorkbook.Sheets("ReportSheetProperties").ListObjects("tbl_ReportProperties")
'
'    For i = LBound(pvtReportMetaData.PvtFields) To UBound(pvtReportMetaData.PvtFields)
'        With pvtReportMetaData.PvtFields(i)
'
'            If .Name = "Values" Then
'                'Special case.  This is only case where orientation is set at a pivot
'                'field level.  Other cases are set at cube field level
'                WriteSingleReportSheetPropertyRecord pvtReportMetaData.SheetName, _
'                    .Name, "PivotieldDataType", "Orientation", .Orientation
'
'
'            ElseIf .Orientation = xlDataField Then
'
'                WriteSingleReportSheetPropertyRecord pvtReportMetaData.SheetName, _
'                   .Name, "PivotieldDataType", "SubtotalName", .SubtotalName
'
'                WriteSingleReportSheetPropertyRecord pvtReportMetaData.SheetName, _
'                   .Name, "PivotieldDataType", "Subtotals", .Subtotals
'
'                WriteSingleReportSheetPropertyRecord pvtReportMetaData.SheetName, _
'                   .Name, "PivotieldDataType", "NumberFormat", .NumberFormat
'
'                WriteSingleReportSheetPropertyRecord pvtReportMetaData.SheetName, _
'                   .Name, "PivotieldDataType", "LayoutSubtotalLocation", .LayoutSubtotalLocation
'
'            Else
'
'                WriteSingleReportSheetPropertyRecord pvtReportMetaData.SheetName, _
'                   .Name, "PivotieldDataType", "LayoutBlankLine", .LayoutBlankLine
'
'                WriteSingleReportSheetPropertyRecord pvtReportMetaData.SheetName, _
'                   .Name, "PivotieldDataType", "LayoutCompactRow", .LayoutCompactRow
'
'                WriteSingleReportSheetPropertyRecord pvtReportMetaData.SheetName, _
'                   .Name, "PivotieldDataType", "LayoutForm", .LayoutForm
'
'                WriteSingleReportSheetPropertyRecord pvtReportMetaData.SheetName, _
'                   .Name, "PivotieldDataType", "LayoutPageBreak", .LayoutPageBreak
'
'                WriteSingleReportSheetPropertyRecord pvtReportMetaData.SheetName, _
'                   .Name, "PivotieldDataType", "RepeatLabels", .RepeatLabels
'
'            End If
'
'        End With
'    Next i
'
'
'End Sub
'
'Function ReadPivotReportSheetNameBasedOnReportHeading(ByVal sSheetHeading As String) As String
''Returns corresponding SheetName for sheetheading as per table on sheet ReportSheetProperties
'
'    Dim sEval As String
'
'    sEval = "=XLOOKUP(" & _
'        """SheetDataTypeSheetHeading" & sSheetHeading & """" & _
'        ", tbl_ReportProperties[DataType] & tbl_ReportProperties[Property] & tbl_ReportProperties[Value]," & _
'        "tbl_ReportProperties[SheetName])"
'
'     ReadPivotReportSheetNameBasedOnReportHeading = Application.Evaluate(sEval)
'
'
'End Function
'
'
'Function ReadPivotTableNameBasedOnSheetName(ByVal sSheetName As String) As String
''Returns pivot table name based on sheet name on sheet ReportSheetProperties
'
'    Dim sEval As String
'
'    sEval = "=XLOOKUP(" & _
'                    """" & sSheetName & "PivotTableDataType""," & _
'                    "tbl_ReportProperties[SheetName] & tbl_ReportProperties[DataType]," & _
'                    "tbl_ReportProperties[Name])"
'
'     ReadPivotTableNameBasedOnSheetName = Application.Evaluate(sEval)
'
'
'End Function
'
'
'
'
'
'Function ReadSinglePivotReportPropertyValue(ByVal sSheetName As String, _
'    ByVal sName As String, _
'    ByVal sDataType As String, _
'    ByVal sProperty As String)
''Reads Value field for given paramaters on sheet "ReportSheetProperties"
'
'    Dim sEvalStr As String
'
'    sEvalStr = "=XLOOKUP(""" & sSheetName & sName & sDataType & sProperty & """, " & _
'        "tbl_ReportProperties[SheetName]&tbl_ReportProperties[Name]& tbl_ReportProperties[DataType]&tbl_ReportProperties[Property]," & _
'        "tbl_ReportProperties[Value])"
'
'    ReadSinglePivotReportPropertyValue = Evaluate(sEvalStr)
'
'End Function
'
'
'Function ReadPivotReportMetaData(ByVal sSheetHeading As String) As TypePivotReport
''Reads the pivot report metadata from the sheet ReportSheetProperties'
'
'
'
''    Dim iSheetNameColumnNumber As Integer
''    Dim iNameColumnNumber As Integer
''    Dim iDataTypeColumnNumber As Integer
''    Dim lo As ListObject
'    Dim pvtReportMetaData As TypePivotReport
'    Dim sSheetName As String
'    Dim i As Double
'    Dim sProperty As String
'    Dim sValue As String
'    Dim vPivotCubeFieldArray As Variant
'    Dim vPivotFieldArray As Variant
'
'    sSheetName = ReadPivotReportSheetNameBasedOnReportHeading(sSheetHeading)
'    pvtReportMetaData.SheetName = sSheetName
''    Set lo = ActiveWorkbook.Sheets("ReportSheetProperties").ListObjects("tbl_ReportProperties")
''
''    lo.ShowAutoFilter = True
''    lo.AutoFilter.ShowAllData
''    iSheetNameColumnNumber = WorksheetFunction.Match("SheetName", lo.HeaderRowRange, 0)
''    iNameColumnNumber = WorksheetFunction.Match("Name", lo.HeaderRowRange, 0)
''    iDataTypeColumnNumber = WorksheetFunction.Match("DataType", lo.HeaderRowRange, 0)
''    lo.Range.AutoFilter field:=iSheetNameColumnNumber, Criteria1:=sSheetName
'
'    With pvtReportMetaData.ReportingSheet
'        .Name = sSheetName
'        .SheetCategory = ReadSinglePivotReportPropertyValue(sSheetName, .Name, "SheetDataType", _
'            "SheetCategory")
'        .SheetHeading = ReadSinglePivotReportPropertyValue(sSheetName, .Name, "SheetDataType", _
'            "SheetHeading")
'    End With
'
'    With pvtReportMetaData.PvtTable
'
'    End With
'
'
'
''    'Read the table properties
''    lo.Range.AutoFilter field:=iDataTypeColumnNumber, Criteria1:="PivotTableDataType"
''    pvtReportMetaData.PvtTable.Name = lo.ListColumns("Name").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1)
''    Set pvtReportMetaData.PvtTable.Properties = New Dictionary
''    For i = 1 To lo.ListColumns("SheetName").DataBodyRange.SpecialCells(xlCellTypeVisible).Count
''        sProperty = lo.ListColumns("Property").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(i)
''        sValue = lo.ListColumns("Value").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(i)
''        pvtReportMetaData.PvtTable.Properties.Add sProperty, sValue
''    Next i
''
''    'Read the multiple pivot cube fields properties
''    lo.Range.AutoFilter field:=iDataTypeColumnNumber, Criteria1:="PivotCubeFieldDataType"
''    vPivotCubeFieldArray = WorksheetFunction.Unique(lo.ListColumns("Name").DataBodyRange.SpecialCells(xlCellTypeVisible))
''    ReDim pvtReportMetaData.PvtCubeFields(UBound(vPivotCubeFieldArray, 1) - 1)
''    For j = 0 To UBound(vPivotCubeFieldArray, 1) - 1
''        lo.Range.AutoFilter field:=iNameColumnNumber, Criteria1:=vPivotCubeFieldArray(j + 1, 1)
''        pvtReportMetaData.PvtCubeFields(j).Name = lo.ListColumns("Name").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1)
''        Set pvtReportMetaData.PvtCubeFields(j).Properties = New Dictionary
''        For i = 1 To lo.ListColumns("SheetName").DataBodyRange.SpecialCells(xlCellTypeVisible).Count
''            sProperty = lo.ListColumns("Property").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(i)
''            sValue = lo.ListColumns("Value").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(i)
''            pvtReportMetaData.PvtCubeFields(j).Properties.Add sProperty, sValue
''        Next i
''    Next j
''
''    'Read the multiple pivot fields properties
''    lo.Range.AutoFilter field:=iNameColumnNumber  'Need to remove filter set above
''    lo.Range.AutoFilter field:=iDataTypeColumnNumber, Criteria1:="PivotFieldDataType"
''    vPivotFieldArray = WorksheetFunction.Unique(lo.ListColumns("Name").DataBodyRange.SpecialCells(xlCellTypeVisible))
''    ReDim pvtReportMetaData.PvtFields(UBound(vPivotFieldArray, 1) - 1)
''    For j = 0 To UBound(vPivotFieldArray, 1) - 1
''        lo.Range.AutoFilter field:=iNameColumnNumber, Criteria1:=vPivotFieldArray(j + 1, 1)
''        pvtReportMetaData.PvtFields(j).Name = lo.ListColumns("Name").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1)
''        Set pvtReportMetaData.PvtFields(j).Properties = New Dictionary
''        For i = 1 To lo.ListColumns("SheetName").DataBodyRange.SpecialCells(xlCellTypeVisible).Count
''            sProperty = lo.ListColumns("Property").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(i)
''            sValue = lo.ListColumns("Value").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(i)
''            pvtReportMetaData.PvtFields(j).Properties.Add sProperty, sValue
''        Next i
''    Next j
''
''
''    lo.AutoFilter.ShowAllData
'
'
'    ReadPivotReportMetaData = pvtReportMetaData
'
'
'End Function
