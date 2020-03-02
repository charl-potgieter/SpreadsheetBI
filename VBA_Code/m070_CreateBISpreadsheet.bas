Attribute VB_Name = "m070_CreateBISpreadsheet"
Option Explicit
Option Private Module


Sub CreateParameterSheet(ByRef wkb As Workbook)

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
    
    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 6
    ActiveWindow.FreezePanes = True


End Sub

Sub CreateValidationSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet

    'Create report report fields sheet
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "Validations"
    sht.Range("SheetHeading") = "Validations"
    sht.Range("SheetCategory") = "Setup"
   
    sht.Range("B6") = "Model Measures"
    sht.Range("C6") = "Model Columns"
    sht.Range("6:6").Font.Bold = True
    sht.Range("6:6").HorizontalAlignment = xlCenter
    wkb.Names.Add Name:="val_Measures", RefersTo:="=Validations!$B$7"
    wkb.Names.Add Name:="val_Columns", RefersTo:="=Validations!$C$7"
    
    sht.Range("B:B").ColumnWidth = 40
    sht.Range("C:C").ColumnWidth = 40


    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 6
    ActiveWindow.FreezePanes = True


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

    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 7
    ActiveWindow.FreezePanes = True


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

    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 6
    ActiveWindow.FreezePanes = True

End Sub





Sub CreateReportPropertiesSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject

    'Create report properties sheet
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ReportProperties"
    sht.Range("SheetHeading") = "Report properties"
    sht.Range("SheetCategory") = "Setup"
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:E8"), XlListObjectHasHeaders:=xlYes)
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

    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 6
    ActiveWindow.FreezePanes = True


End Sub


Sub CreateReportFieldSettingsSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject
    Dim sRelativeReferenceOfDataFieldType As String
    Dim sValidationStr As String
    Dim DataVal As FormatCondition

    'Create report report fields sheet
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ReportFieldSettings"
    sht.Range("SheetHeading") = "Report field settings"
    sht.Range("SheetCategory") = "Setup"
    sht.Range("B5") = "Run ""Data model update > Write Measures to columns and sheets"" to refresh validation "
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B7:G15"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_ReportFields"
        .HeaderRowRange.Cells(1) = "Report Name"
        .HeaderRowRange.Cells(2) = "Data Model Field Type"
        .HeaderRowRange.Cells(3) = "Cube Field Name"
        .HeaderRowRange.Cells(4) = "Orientation"
        .HeaderRowRange.Cells(5) = "Format"
        .HeaderRowRange.Cells(6) = "Custom Format"
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 40
    sht.Range("C:C").ColumnWidth = 20
    sht.Range("D:D").ColumnWidth = 40
    sht.Range("E:E").ColumnWidth = 20
    sht.Range("F:F").ColumnWidth = 20
    sht.Range("G:G").ColumnWidth = 20

    'Set cube field validations (cascading depending on field type)
    sRelativeReferenceOfDataFieldType = Replace(lo.ListColumns("Data Model Field Type").DataBodyRange.Cells(1).Address, "$", "")
    sValidationStr = "=INDIRECT(""val_"" & IF(" & sRelativeReferenceOfDataFieldType & " ="""", ""Measure"", " & sRelativeReferenceOfDataFieldType & ") & ""s"")"
    lo.ListColumns("Cube Field Name").DataBodyRange.Validation.Add _
        Type:=xlValidateList, Formula1:=sValidationStr
    
    'Set other validations
    lo.ListColumns("Data Model Field Type").DataBodyRange.Validation.Add _
        Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Measure, Column"


    lo.ListColumns("Orientation").DataBodyRange.Validation.Add _
        Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Row, Column, Filter"

    lo.ListColumns("Format").DataBodyRange.Validation.Add _
        Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Zero Decimals,One Decimal,Two Decimals,Custom"
        
        
    'Set conditionalFormatting
    Set DataVal = lo.ListColumns("Orientation").DataBodyRange.FormatConditions.Add _
        (Type:=xlExpression, _
        Formula1:="=" & sRelativeReferenceOfDataFieldType & " = ""Measure""")
    DataVal.Interior.Color = RGB(0, 0, 0)
    
    Set DataVal = lo.ListColumns("Format").DataBodyRange.FormatConditions.Add _
        (Type:=xlExpression, _
        Formula1:="=" & sRelativeReferenceOfDataFieldType & " = ""Column""")
    DataVal.Interior.Color = RGB(0, 0, 0)
    
    Set DataVal = lo.ListColumns("Custom Format").DataBodyRange.FormatConditions.Add _
        (Type:=xlExpression, _
        Formula1:="=" & sRelativeReferenceOfDataFieldType & " = ""Column""")
    DataVal.Interior.Color = RGB(0, 0, 0)
        
        
    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 7
    ActiveWindow.FreezePanes = True

End Sub



Sub CreateModelMeasuresSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject

    
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ModelMeasures"
    sht.Range("SheetHeading") = "Data model measures"
    sht.Range("SheetCategory") = "Setup"
   
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:F7"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_ModelMeasures"
        .HeaderRowRange.Cells(1) = "Name"
        .HeaderRowRange.Cells(2) = "Visible"
        .HeaderRowRange.Cells(3) = "Unique Name"
        .HeaderRowRange.Cells(4) = "DAX Expression"
        .HeaderRowRange.Cells(5) = "Name and Expression"
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 40
    sht.Range("C:C").ColumnWidth = 20
    sht.Range("D:D").ColumnWidth = 40
    sht.Range("E:E").ColumnWidth = 80
    sht.Range("F:F").ColumnWidth = 80

    With lo.DataBodyRange
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With

    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 6
    ActiveWindow.FreezePanes = True

End Sub


Sub CreateModelColumnsSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject

   
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ModelColumns"
    sht.Range("SheetHeading") = "Data model columns"
    sht.Range("SheetCategory") = "Setup"
    sht.Range("B4") = "Includes calculated columns"
   
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:F7"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_ModelColumns"
        .HeaderRowRange.Cells(1) = "Name"
        .HeaderRowRange.Cells(2) = "Table Name"
        .HeaderRowRange.Cells(3) = "Unique Name"
        .HeaderRowRange.Cells(4) = "Visible"
        .HeaderRowRange.Cells(5) = "Is calculated column"
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 30
    sht.Range("C:C").ColumnWidth = 30
    sht.Range("D:D").ColumnWidth = 50
    sht.Range("E:E").ColumnWidth = 20
    sht.Range("F:F").ColumnWidth = 20

    With lo.DataBodyRange
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With

    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 6
    ActiveWindow.FreezePanes = True


End Sub






Sub CreateModelCalculatedColumnsSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject

    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ModelCalcColumns"
    sht.Range("SheetHeading") = "Data model calculated columns"
    sht.Range("SheetCategory") = "Setup"
   
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:D7"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_ModelCalcColumns"
        .HeaderRowRange.Cells(1) = "Name"
        .HeaderRowRange.Cells(2) = "Table Name"
        .HeaderRowRange.Cells(3) = "Expression"
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 30
    sht.Range("C:C").ColumnWidth = 30
    sht.Range("D:D").ColumnWidth = 50

    With lo.DataBodyRange
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With

    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 6
    ActiveWindow.FreezePanes = True


End Sub



Sub CreateModelRelationshipsSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject

    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ModelRelationships"
    sht.Range("SheetHeading") = "Data model relationships"
    sht.Range("SheetCategory") = "Setup"
   
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:F7"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_ModelRelationships"
        .HeaderRowRange.Cells(1) = "Primary Key Table"
        .HeaderRowRange.Cells(2) = "Primary Key Column"
        .HeaderRowRange.Cells(3) = "Foreign Key Table"
        .HeaderRowRange.Cells(4) = "Foreign Key Column"
        .HeaderRowRange.Cells(5) = "Active"
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 40
    sht.Range("C:C").ColumnWidth = 40
    sht.Range("D:D").ColumnWidth = 40
    sht.Range("E:E").ColumnWidth = 40
    sht.Range("F:F").ColumnWidth = 20

    With lo.DataBodyRange
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With

    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 6
    ActiveWindow.FreezePanes = True


End Sub

