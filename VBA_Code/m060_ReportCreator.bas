Attribute VB_Name = "m060_ReportCreator"
Option Explicit

Sub AAATestCreateReport()


    CreatePivotTable "Test Report 1"


End Sub


Sub CreatePivotTable(ByVal sReportName As String)

    Dim pvt As PivotTable
    Dim loPivotTableSettings As ListObject
    Dim loPivotFieldSettings As ListObject
    

    Set pvt = ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
        ActiveWorkbook.Connections("ThisWorkbookDataModel"), Version:=6). _
        CreatePivotTable(ActiveCell)
    
 

    
    
    'Set pivot table properties
    Set loPivotFieldSettings = ActiveWorkbook.Sheets("ReportFields").ListObjects("tbl_ReportFields")
    
    'Set pivot field properties
    Set loPivotFieldSettings = ActiveWorkbook.Sheets("ReportFields").ListObjects("tbl_ReportFields")
    SetPivotField pvt, loPivotFieldSettings, sReportName
    
    
End Sub


Sub CustomisePivotTable()




End Sub




Sub SetPivotField(ByRef pvt As PivotTable, ByRef loPivotFieldSettings As ListObject, sReportName As String)

    Dim i As Double
    Dim sCubeFieldName As String

    With loPivotFieldSettings
        For i = 1 To .DataBodyRange.Rows.Count
    
            If .ListColumns("Report Name").DataBodyRange.Cells(i) = sReportName Then
            
                sCubeFieldName = .ListColumns("Cube Field Name").DataBodyRange.Cells(i)
                
                'Set field on either data, row of colum
                Select Case .ListColumns("Orientation").DataBodyRange.Cells(i)
                    Case "Data"
                        pvt.CubeFields(sCubeFieldName).Orientation = xlDataField
                    Case "Row"
                        pvt.CubeFields(sCubeFieldName).Orientation = xlRowField
                    Case "Column"
                        pvt.CubeFields(sCubeFieldName).Orientation = xlColumnField
                End Select
                
                'Format field
                Select Case .ListColumns("Format").DataBodyRange.Cells(i)
                    Case "Zero Decimals"
                        pvt.PivotFields(sCubeFieldName).NumberFormat = "#,##0_);(#,##0);-??"
                    Case "One Decimal"
                        pvt.PivotFields(sCubeFieldName).NumberFormat = "#,##0.0_);(#,##0.0);-??"
                    Case "Two Decimals"
                        pvt.PivotFields(sCubeFieldName).NumberFormat = "#,##0.00_);(#,##0.00);-??"
                End Select
                
            End If

        Next i
    
    End With

End Sub





















