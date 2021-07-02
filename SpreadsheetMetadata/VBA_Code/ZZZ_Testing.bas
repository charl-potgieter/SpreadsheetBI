Attribute VB_Name = "ZZZ_Testing"
Option Explicit

Sub TestVariantArray()

    Dim wkb As Workbook
    Dim rng As Range
    Dim v As Variant
    Dim colProperty As Variant
    Dim colValue As Variant
    Dim unq As Variant
    Dim rv As Variant
    
    Set wkb = Application.Workbooks("PivotReportExample.xlsm")
    Set rng = wkb.Sheets("ReportPRoperties").ListObjects("tbl_ReportProperties").DataBodyRange
    v = rng
    colProperty = rng.Columns(4)
    colValue = rng.Columns(5)
    
    rv = WorksheetFunction.Xlookup("Category", colProperty, colValue)

End Sub

