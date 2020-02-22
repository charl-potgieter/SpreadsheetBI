Attribute VB_Name = "m060_ReportCreator"
Option Explicit
Option Private Module


Function ReportSettingsAreValid() As Boolean

    MsgBox ("was here")


End Function



Sub CreatePivotTable(ByVal sReportName As String, ByRef pvt As PivotTable)

    Dim loPivotTableSettings As ListObject
    Dim loPivotFieldSettings As ListObject


    Set pvt = ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
        ActiveWorkbook.Connections("ThisWorkbookDataModel"), Version:=6). _
        CreatePivotTable(ActiveCell)
    
    
End Sub


Sub CustomisePivotTable(ByVal sReportName As String)

'!!!!!!!!!!!!!!!!!!!!! TO BUILD


End Sub




Sub SetPivotFields(ByRef pvt As PivotTable, ByRef loPivotFieldSettings As ListObject, sReportName As String)

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





















