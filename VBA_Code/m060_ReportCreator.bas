Attribute VB_Name = "m060_ReportCreator"
Option Explicit
Option Private Module


Function ReportSettingsAreValid() As Boolean

    MsgBox ("Need to still build Settings check")
    
    ReportSettingsAreValid = True


End Function



Sub CreatePivotTable(ByVal sSheetName As String, ByVal sReportName As String, ByVal sReportCategory As String, ByRef pvt As PivotTable)

    Dim loPivotTableSettings As ListObject
    Dim loPivotFieldSettings As ListObject
    Dim sht As Worksheet

    If SheetExists(ActiveWorkbook, sSheetName) Then
        MsgBox ("deleting sheet, need to give user a choice")
        ActiveWorkbook.Sheets(sSheetName).Delete
    End If

    Set sht = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    sht.Name = sSheetName
    FormatSheet sht
    sht.Range("SheetHeading") = sReportName
    sht.Range("SheetCategory") = sReportCategory
    
    Set pvt = ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
        ActiveWorkbook.Connections("ThisWorkbookDataModel"), Version:=6). _
        CreatePivotTable(sht.Range("B10"))
    
    
End Sub


Sub CustomisePivotTable(ByVal sReportName As String)

'!!!!!!!!!!!!!!!!!!!!! TO BUILD


End Sub




Sub SetPivotFields(ByRef pvt As PivotTable, ByVal sCubeFieldName As String, ByVal sOrientation As String)

    Dim i As Double

                
    'Set field on either data, row of colum
    Select Case sOrientation
        Case "Data"
            pvt.CubeFields(sCubeFieldName).Orientation = xlDataField
        Case "Row"
            pvt.CubeFields(sCubeFieldName).Orientation = xlRowField
        Case "Column"
            pvt.CubeFields(sCubeFieldName).Orientation = xlColumnField
    End Select
    
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





















