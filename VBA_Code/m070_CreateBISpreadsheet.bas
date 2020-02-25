Attribute VB_Name = "m070_CreateBISpreadsheet"
Option Explicit
Option Private Module


Sub CreateParamaterSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject
    
    Set sht = wkb.Sheets(1)
    FormatSheet sht
    sht.Name = "Parameters"
    sht.Range("SheetHeading") = "Parameters"
    sht.Range("SheetCategory") = "Setup"
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:C12"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_Parameters"
        .HeaderRowRange.Cells(1) = "Parameter"
        .HeaderRowRange.Cells(2) = "Value"
        .ListColumns("Parameter").DataBodyRange.Cells(1) = "Date_Start"
        .ListColumns("Parameter").DataBodyRange.Cells(2) = "Date_End"
        .ListColumns("Value").DataBodyRange.Cells(1) = DateSerial(2018, 1, 1)
        .ListColumns("Value").DataBodyRange.Cells(2) = DateSerial(2020, 12, 31)
        .ListColumns("Value").DataBodyRange.Cells(1).NumberFormat = "dd-mmm-yy"
        .ListColumns("Value").DataBodyRange.Cells(2).NumberFormat = "dd-mmm-yy"
        .HeaderRowRange.RowHeight = .HeaderRowRange.RowHeight * 2
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 30
    sht.Range("C:C").ColumnWidth = 60
    

End Sub



Sub CreateReportListSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject

    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ReportList"
    sht.Range("SheetHeading") = "Report List"
    sht.Range("SheetCategory") = "Setup"
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B7:F10"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_ReportList"
        .HeaderRowRange.Cells(1) = "Report Name"
        .HeaderRowRange.Cells(2) = "Sheet Name"
        .HeaderRowRange.Cells(3) = "Report Category"
        .HeaderRowRange.Cells(4) = "Run with table refresh"
        .HeaderRowRange.Cells(5) = "Run without table refresh"
        .HeaderRowRange.RowHeight = .HeaderRowRange.RowHeight * 2
        .ListColumns("Run with table refresh").DataBodyRange.HorizontalAlignment = xlCenter
        .ListColumns("Run without table refresh").DataBodyRange.HorizontalAlignment = xlCenter
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 60
    sht.Range("C:C").ColumnWidth = 30
    sht.Range("D:D").ColumnWidth = 30
    sht.Range("B5") = "Clear data from non-dependent tables (mark with X)"
    
    
    
    sht.Names.Add Name:="ClearData", RefersTo:="=$F$5"

    With sht.Range("ClearData")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(217, 225, 242)
    End With
    SetOuterBorders sht.Range("ClearData")

End Sub


Sub CreateQueriesPerReportSheet(ByRef wkb As Workbook)
    
    Dim sht As Worksheet
    Dim lo As ListObject
    
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "QueriesPerReport"
    sht.Range("SheetHeading") = "Queries per report"
    sht.Range("SheetCategory") = "Setup"
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:D8"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_QueriesPerReport"
        .HeaderRowRange.Cells(1) = "Report Name"
        .HeaderRowRange.Cells(2) = "Report selected for run and query refresh"
        .HeaderRowRange.Cells(3) = "Query Name"
        .HeaderRowRange.RowHeight = .HeaderRowRange.RowHeight * 2
        .ListColumns("Report selected for run and query refresh").DataBodyRange.Formula = _
             "=(COUNTIFS(tbl_ReportList[Report Name], [@[Report Name]], tbl_ReportList[Run with table refresh], ""*"")) > 0"
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 50
    sht.Range("C:C").ColumnWidth = 30
    sht.Range("D:D").ColumnWidth = 50

End Sub





Sub CreateReportPropertiesSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject
    Dim sFormulaStr As String

    'Create report properties sheet
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ReportProperties"
    sht.Range("SheetHeading") = "Report properties"
    sht.Range("SheetCategory") = "Setup"
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:D8"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_ReportProperties"
        .HeaderRowRange.Cells(1) = "Report Name"
        .HeaderRowRange.Cells(2) = "AutoFit"
        .HeaderRowRange.Cells(3) = "Total Rows"
        .HeaderRowRange.Cells(4) = "Total Columns"
        .HeaderRowRange.RowHeight = .HeaderRowRange.RowHeight * 2
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 50
    sht.Range("C:C").ColumnWidth = 20
    sht.Range("D:D").ColumnWidth = 20
    sht.Range("E:E").ColumnWidth = 20


    'Create report report fields sheet
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ReportFields"
    sht.Range("SheetHeading") = "Report fields"
    sht.Range("SheetCategory") = "Setup"
    sht.Range("B5") = "Run ""Data model update > Update report field validation"" to refresh "
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B7:F15"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_ReportFields"
        .HeaderRowRange.Cells(1) = "Report Name"
        .HeaderRowRange.Cells(2) = "Cube Field Name"
        .HeaderRowRange.Cells(3) = "Orientation"
        .HeaderRowRange.Cells(4) = "Format"
        .HeaderRowRange.Cells(5) = "Custom Format"
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 40
    sht.Range("C:C").ColumnWidth = 40
    sht.Range("D:D").ColumnWidth = 20
    sht.Range("E:E").ColumnWidth = 20
    sht.Range("F:F").ColumnWidth = 20

    lo.ListColumns("Orientation").DataBodyRange.Validation.Add _
        Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Row, Column, Filter, Data"


    lo.ListColumns("Format").DataBodyRange.Validation.Add _
        Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Zero Decimals,One Decimal,Two Decimals,Custom"

    'Create check to ensure that measure and fields have correct orientation
    sFormulaStr = "=IF(" & Chr(10) & _
        "       OR(" & Chr(10) & _
        "                 COUNTIFS(tbl_ReportFields[Cube Field Name], ""*measure*"", tbl_ReportFields[Orientation], ""<>Data"") >0," & Chr(10) & _
        "                 COUNTIFS(tbl_ReportFields[Cube Field Name], ""<>*measure*"", tbl_ReportFields[Orientation], ""Data"") >0" & Chr(10) & _
        "               )," & Chr(10) & _
        "            ""Orientation for measures need to be set as 'data', all other fields must be either row, column or filter""," & Chr(10) & _
        "            ""Ok""" & Chr(10) & _
        "      )"
        
    With sht.Range("C3")
        .Formula = sFormulaStr
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$C$3<>""OK"""
        .FormatConditions(1).Font.Bold = True
        .FormatConditions(1).Font.Color = RGB(255, 0, 0)
        '.FormatConditions(1).Font.Color = -16776961
    End With


End Sub
