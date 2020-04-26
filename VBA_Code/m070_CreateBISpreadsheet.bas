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
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:F9"), XlListObjectHasHeaders:=xlYes)
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
    

    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 7
    ActiveWindow.FreezePanes = True


End Sub


Sub CreateDataLoadQueriesPerReport(ByRef wkb As Workbook)
    
    Dim sht As Worksheet
    Dim lo As ListObject
    
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "DataLoadQueriesPerReport"
    sht.Range("SheetHeading") = "Data load queries per report"
    sht.Range("SheetCategory") = "Setup"
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:C8"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_DataLoadQueriesPerReport"
        .HeaderRowRange.Cells(1) = "Report Name"
        .HeaderRowRange.Cells(2) = "Data Load Query Name"
        .HeaderRowRange.RowHeight = .HeaderRowRange.RowHeight * 2
        
        FormatTable lo
        .ListColumns("Report Name").DataBodyRange.EntireColumn.ColumnWidth = 50
        .ListColumns("Data Load Query Name").DataBodyRange.EntireColumn.ColumnWidth = 50
        
    End With
    


    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 6
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




Sub CreateTableGeneratorSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject

    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "TableGenerator"
    sht.Range("SheetHeading") = "Table Generator"
    sht.Range("SheetCategory") = "Setup"
   
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B11:F12"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_TableGenerator"
        .HeaderRowRange.Cells(1) = "Column_1"
        .HeaderRowRange.Cells(2) = "Column_2"
        .HeaderRowRange.Cells(3) = "Column_3"
        .HeaderRowRange.Cells(4) = "Column_4"
        .HeaderRowRange.Cells(5) = "Column_5"
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 20
    sht.Range("C:C").ColumnWidth = 20
    sht.Range("D:D").ColumnWidth = 20
    sht.Range("E:E").ColumnWidth = 20
    sht.Range("F:F").ColumnWidth = 20

    With lo.DataBodyRange
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With

    'Add various formatted text to the sheet
    sht.Range("B5") = "Generates a power query with hardcoded values and field types as below, using the GeneratePowerQuery code"
    sht.Range("B7") = "Query Name"
    sht.Range("C7") = "TestTable"
    sht.Range("B7").Font.Bold = True
    sht.Range("C7,B9:E9").Interior.Color = RGB(242, 242, 242)
    sht.Range("C7,B9:E9").Font.Color = RGB(0, 112, 192)
    sht.Range("C7,B9:E9").HorizontalAlignment = xlCenter
    sht.Range("B9:E9") = "text"
    
    'Add data validation for field types
    sht.Range("B9:F9").Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:="any,binary,date,datetime,datetimezone,duration,function,function,Int64.Type,list,list," & _
            "logical,none,null,number,record,record,table,table,text,time,type"
  

    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 11
    ActiveWindow.FreezePanes = True


End Sub
