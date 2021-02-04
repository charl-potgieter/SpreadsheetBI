Attribute VB_Name = "m070_ReadReportMetaData"
Option Explicit
Option Private Module

Sub CreateReportMetaDataSheets()

    Dim shtReportSheetProperties As Worksheet
    Dim shtPivotTableProperties As Worksheet
    Dim shtPivotFieldProperties As Worksheet
    Dim lo As ListObject


    'Create sheet for report property capture
    If Not SheetExists(ActiveWorkbook, "ReportSheetProperties") Then
        Set shtReportSheetProperties = InsertFormattedSheetIntoActiveWorkbook
        With shtReportSheetProperties
            .Name = "ReportSheetProperties"
            .Range("SheetCategory").Value = "Report metadata"
            .Range("SheetHeading").Value = "Report sheet properties"
            
            'Create table headings
            .Range("B6") = "Sheet Name|Sheet Category"
            .Range("B6").TextToColumns DataType:=xlDelimited, Other:=True, OtherChar:="|"
            
            'Ensure at least one row in table
            .Range("B7") = "Empty Table"
            
            Set lo = .ListObjects.Add(xlSrcRange, Range("$B$6").CurrentRegion, , xlYes)
        End With
        lo.Name = "tbl_ReportSheetProperties"
        FormatTable lo
    End If

    
    'Create sheet for pivot table property capture
    If Not SheetExists(ActiveWorkbook, "PvtTableProperties") Then
        Set shtPivotTableProperties = InsertFormattedSheetIntoActiveWorkbook
        With shtPivotTableProperties
            .Name = "PvtTableProperties"
            .Range("SheetCategory").Value = "Report metadata"
            .Range("SheetHeading").Value = "Pivot table properties"
            
            'Create table headings
            .Range("B6") = "Sheet Name|Pivot Table Name|Auto Fit|Total Rows|Total Columns|Display Expand Buttons|Display Field Headers"
            .Range("B6").TextToColumns DataType:=xlDelimited, Other:=True, OtherChar:="|"
            
            'Ensure at least one row in table
            .Range("B7") = "Empty Table"
            
            Set lo = .ListObjects.Add(xlSrcRange, Range("$B$6").CurrentRegion, , xlYes)
        End With
        lo.Name = "tbl_PvtTableProperties"
        FormatTable lo
    End If
    
    'Create sheet for pivot field property capture
    If Not SheetExists(ActiveWorkbook, "PvtFieldProperties") Then
        Set shtPivotFieldProperties = InsertFormattedSheetIntoActiveWorkbook
        With shtPivotFieldProperties
            .Name = "PvtFieldProperties"
            .Range("SheetCategory").Value = "Report metadata"
            .Range("SheetHeading").Value = "Pivot field properties"
            
            'Create table headings
            .Range("B6") = "Sheet Name|Pivot Table Name|Data Model Field Type|Cube Field Name|Orientation|Format|Custom Format|Subtotal" & _
                "|Subtotal at top|Blank line between items|Filter Type|Filter Values|Collapse field values"
            .Range("B6").TextToColumns DataType:=xlDelimited, Other:=True, OtherChar:="|"
            
            'Ensure at least one row in table
            .Range("B7") = "Empty Table"
            
            Set lo = .ListObjects.Add(xlSrcRange, Range("$B$6").CurrentRegion, , xlYes)
        End With
        lo.Name = "tbl_PvtTableProperties"
        FormatTable lo
    End If
            
    
    'Set sheet order
    ActiveWorkbook.Sheets("ReportSheetProperties").Move After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    ActiveWorkbook.Sheets("PvtTableProperties").Move After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    ActiveWorkbook.Sheets("PvtFieldProperties").Move After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    
    

End Sub




