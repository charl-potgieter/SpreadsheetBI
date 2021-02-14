Attribute VB_Name = "m006_DataAccess"
'Option Explicit
'Option Private Module

Sub SavePivotReportMetaData(ByRef pvtReportMetaData() As TypePivotReport)
'Saves pivot report metadata on sheeet in active workbook

    Dim i As Integer

    If Not SheetExists(ActiveWorkbook, "ReportSheetProperties") Then
        CreateReportPropertiesSheet ActiveWorkbook
    End If

    For i = LBound(pvtReportMetaData) To UBound(pvtReportMetaData)
        DeleteAnyPreExistingData pvtReportMetaData(i)
        WriteSheetProperties pvtReportMetaData(i)
        WritePivotTableProperties pvtReportMetaData(i)
        WritePivotCubeFieldProperties pvtReportMetaData(i)
        WritePivotFieldProperties pvtReportMetaData(i)
    Next i

    FormatReportSheetProperties

End Sub


Function PivotReportSheetNameBasedOnReportHeading(ByVal sSheetHeading As String) As String
'Returns corresponding SheetName for sheetheading as per table on sheet ReportSheetProperties

    Dim sEval As String

    sEval = "=XLOOKUP(" & _
        """SheetDataTypeSheetHeading" & sSheetHeading & """" & _
        ", tbl_ReportProperties[DataType] & tbl_ReportProperties[Property] & tbl_ReportProperties[Value]," & _
        "tbl_ReportProperties[SheetName])"

     PivotReportSheetNameBasedOnReportHeading = Application.Evaluate(sEval)


End Function



Function ReadPivotReportMetaData(ByVal sSheetName As String) As TypePivotReport
'Reads the pivot report metadata from the sheet ReportSheetProperties'


    Dim pvtReportMetaData As TypePivotReport
    Dim iSheetNameColumnNumber As Integer
    Dim iNameColumnNumber As Integer
    Dim iDataTypeColumnNumber As Integer
    Dim lo As ListObject
    Dim i As Double
    Dim j As Double
    Dim sProperty As String
    Dim sValue As String
    Dim vPivotCubeFieldArray As Variant
    Dim vPivotFieldArray As Variant
    

    pvtReportMetaData.SheetName = sSheetName
    Set lo = ActiveWorkbook.Sheets("ReportSheetProperties").ListObjects("tbl_ReportProperties")

    lo.ShowAutoFilter = True
    lo.AutoFilter.ShowAllData
    iSheetNameColumnNumber = WorksheetFunction.Match("SheetName", lo.HeaderRowRange, 0)
    iNameColumnNumber = WorksheetFunction.Match("Name", lo.HeaderRowRange, 0)
    iDataTypeColumnNumber = WorksheetFunction.Match("DataType", lo.HeaderRowRange, 0)
    lo.Range.AutoFilter field:=iSheetNameColumnNumber, Criteria1:=sSheetName

    'Read the Sheet Properties
    lo.Range.AutoFilter field:=iDataTypeColumnNumber, Criteria1:="SheetDataType"
    pvtReportMetaData.ReportingSheet.Name = lo.ListColumns("Name").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1)
    Set pvtReportMetaData.ReportingSheet.Properties = New Dictionary
    For i = 1 To lo.ListColumns("SheetName").DataBodyRange.SpecialCells(xlCellTypeVisible).Count
        sProperty = lo.ListColumns("Property").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(i)
        sValue = lo.ListColumns("Value").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(i)
        pvtReportMetaData.ReportingSheet.Properties.Add sProperty, sValue
    Next i

    'Read the table properties
    lo.Range.AutoFilter field:=iDataTypeColumnNumber, Criteria1:="PivotTableDataType"
    pvtReportMetaData.PvtTable.Name = lo.ListColumns("Name").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1)
    Set pvtReportMetaData.PvtTable.Properties = New Dictionary
    For i = 1 To lo.ListColumns("SheetName").DataBodyRange.SpecialCells(xlCellTypeVisible).Count
        sProperty = lo.ListColumns("Property").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(i)
        sValue = lo.ListColumns("Value").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(i)
        pvtReportMetaData.PvtTable.Properties.Add sProperty, sValue
    Next i

    'Read the multiple pivot cube fields properties
    lo.Range.AutoFilter field:=iDataTypeColumnNumber, Criteria1:="PivotCubeFieldDataType"
    vPivotCubeFieldArray = WorksheetFunction.Unique(lo.ListColumns("Name").DataBodyRange.SpecialCells(xlCellTypeVisible))
    ReDim pvtReportMetaData.PvtCubeFields(UBound(vPivotCubeFieldArray, 1) - 1)
    For j = 0 To UBound(vPivotCubeFieldArray, 1) - 1
        lo.Range.AutoFilter field:=iNameColumnNumber, Criteria1:=vPivotCubeFieldArray(j + 1, 1)
        pvtReportMetaData.PvtCubeFields(j).Name = lo.ListColumns("Name").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1)
        Set pvtReportMetaData.PvtCubeFields(j).Properties = New Dictionary
        For i = 1 To lo.ListColumns("SheetName").DataBodyRange.SpecialCells(xlCellTypeVisible).Count
            sProperty = lo.ListColumns("Property").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(i)
            sValue = lo.ListColumns("Value").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(i)
            pvtReportMetaData.PvtCubeFields(j).Properties.Add sProperty, sValue
        Next i
    Next j

    'Read the multiple pivot fields properties
    lo.Range.AutoFilter field:=iNameColumnNumber  'Need to remove filter set above
    lo.Range.AutoFilter field:=iDataTypeColumnNumber, Criteria1:="PivotFieldDataType"
    vPivotFieldArray = WorksheetFunction.Unique(lo.ListColumns("Name").DataBodyRange.SpecialCells(xlCellTypeVisible))
    ReDim pvtReportMetaData.PvtFields(UBound(vPivotFieldArray, 1) - 1)
    For j = 0 To UBound(vPivotFieldArray, 1) - 1
        lo.Range.AutoFilter field:=iNameColumnNumber, Criteria1:=vPivotFieldArray(j + 1, 1)
        pvtReportMetaData.PvtFields(j).Name = lo.ListColumns("Name").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1)
        Set pvtReportMetaData.PvtFields(j).Properties = New Dictionary
        For i = 1 To lo.ListColumns("SheetName").DataBodyRange.SpecialCells(xlCellTypeVisible).Count
            sProperty = lo.ListColumns("Property").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(i)
            sValue = lo.ListColumns("Value").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(i)
            pvtReportMetaData.PvtFields(j).Properties.Add sProperty, sValue
        Next i
    Next j


    lo.AutoFilter.ShowAllData


    ReadPivotReportMetaData = pvtReportMetaData


End Function



Sub CreateReportPropertiesSheet(ByRef wkb As Workbook)

    Dim sHeaders As String
    Dim shtPivotTableProperties As Worksheet
    Dim lo As ListObject

    Set shtPivotTableProperties = InsertFormattedSheetIntoActiveWorkbook
    With shtPivotTableProperties
        .Name = "ReportSheetProperties"
        .Range("SheetCategory").Value = "Report metadata"
        .Range("SheetHeading").Value = "Report Sheet Properties"

        'Create table
        .Range("B6") = "SheetName"
        .Range("C6") = "Name"
        .Range("D6") = "DataType"
        .Range("E6") = "Property"
        .Range("F6") = "Value"
        Set lo = .ListObjects.Add(xlSrcRange, Range("$B$6").CurrentRegion, , xlYes)

    End With
    lo.Name = "tbl_ReportProperties"
    FormatTable lo


End Sub

Sub FormatReportSheetProperties()

    ActiveWorkbook.Sheets("ReportSheetProperties").Activate

    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 6
        .FreezePanes = True
    End With

    With ActiveWorkbook.Sheets("ReportSheetProperties")
        .ListObjects("tbl_ReportProperties"). _
            DataBodyRange.HorizontalAlignment = xlLeft
        Cells.EntireColumn.AutoFit
    End With


End Sub



Sub DeleteAnyPreExistingData(ByRef pvtReportMetaData As TypePivotReport)
'Delete any info if it exist for this sheet name

    Dim lo As ListObject
    Dim iSheetNameColumnNumber As Integer

    Set lo = ActiveWorkbook.Worksheets("ReportSheetProperties").ListObjects("tbl_ReportProperties")
    iSheetNameColumnNumber = WorksheetFunction.Match("SheetName", lo.HeaderRowRange, 0)

    lo.Range.AutoFilter field:=iSheetNameColumnNumber, _
        Criteria1:=pvtReportMetaData.SheetName
    On Error Resume Next
    lo.DataBodyRange.SpecialCells(xlCellTypeVisible).EntireRow.Delete (xlShiftUp)
    On Error GoTo 0
    lo.AutoFilter.ShowAllData


End Sub

Sub WriteSheetProperties(ByRef pvtReportMetaData As TypePivotReport)

    Dim key As Variant
    Dim lo As ListObject
    Dim i As Double

    Set lo = ActiveWorkbook.Sheets("ReportSheetProperties").ListObjects("tbl_ReportProperties")

    With pvtReportMetaData.ReportingSheet
        For Each key In .Properties.Keys
            AddOneRowToListObject lo
            i = lo.DataBodyRange.Rows.Count
            lo.ListColumns("SheetName").DataBodyRange.Cells(i) = pvtReportMetaData.SheetName
            lo.ListColumns("Name").DataBodyRange.Cells(i) = .Name
            lo.ListColumns("DataType").DataBodyRange.Cells(i) = "SheetDataType"
            lo.ListColumns("Property").DataBodyRange.Cells(i) = key
            lo.ListColumns("Value").DataBodyRange.Cells(i) = .Properties.item(key)
        Next key
    End With

End Sub


Sub WritePivotTableProperties(ByRef pvtReportMetaData As TypePivotReport)

    Dim key As Variant
    Dim lo As ListObject
    Dim i As Double

    Set lo = ActiveWorkbook.Sheets("ReportSheetProperties").ListObjects("tbl_ReportProperties")

    With pvtReportMetaData.PvtTable
        For Each key In .Properties.Keys
            AddOneRowToListObject lo
            i = lo.DataBodyRange.Rows.Count
            lo.ListColumns("SheetName").DataBodyRange.Cells(i) = pvtReportMetaData.SheetName
            lo.ListColumns("Name").DataBodyRange.Cells(i) = .Name
            lo.ListColumns("DataType").DataBodyRange.Cells(i) = "PivotTableDataType"
            lo.ListColumns("Property").DataBodyRange.Cells(i) = key
            lo.ListColumns("Value").DataBodyRange.Cells(i) = .Properties.item(key)
        Next key
    End With

End Sub

Sub WritePivotCubeFieldProperties(ByRef pvtReportMetaData As TypePivotReport)

    Dim key As Variant
    Dim lo As ListObject
    Dim i As Double
    Dim j As Double

    Set lo = ActiveWorkbook.Sheets("ReportSheetProperties").ListObjects("tbl_ReportProperties")

    For j = LBound(pvtReportMetaData.PvtCubeFields) To UBound(pvtReportMetaData.PvtCubeFields)
        With pvtReportMetaData.PvtCubeFields(j)

            For Each key In .Properties.Keys
                AddOneRowToListObject lo
                i = lo.DataBodyRange.Rows.Count
                lo.ListColumns("SheetName").DataBodyRange.Cells(i) = pvtReportMetaData.SheetName
                lo.ListColumns("Name").DataBodyRange.Cells(i) = .Name
                lo.ListColumns("DataType").DataBodyRange.Cells(i) = "PivotCubeFieldDataType"
                lo.ListColumns("Property").DataBodyRange.Cells(i) = key
                lo.ListColumns("Value").DataBodyRange.Cells(i) = .Properties.item(key)
            Next key

        End With
    Next j

End Sub



Sub WritePivotFieldProperties(ByRef pvtReportMetaData As TypePivotReport)

    Dim key As Variant
    Dim lo As ListObject
    Dim i As Double
    Dim j As Double

    Set lo = ActiveWorkbook.Sheets("ReportSheetProperties").ListObjects("tbl_ReportProperties")

    For j = LBound(pvtReportMetaData.PvtFields) To UBound(pvtReportMetaData.PvtFields)
        With pvtReportMetaData.PvtFields(j)

            For Each key In .Properties.Keys
                AddOneRowToListObject lo
                i = lo.DataBodyRange.Rows.Count
                lo.ListColumns("SheetName").DataBodyRange.Cells(i) = pvtReportMetaData.SheetName
                lo.ListColumns("Name").DataBodyRange.Cells(i) = .Name
                lo.ListColumns("DataType").DataBodyRange.Cells(i) = "PivotFieldDataType"
                lo.ListColumns("Property").DataBodyRange.Cells(i) = key
                lo.ListColumns("Value").DataBodyRange.Cells(i) = .Properties.item(key)
            Next key

        End With
    Next j



End Sub
