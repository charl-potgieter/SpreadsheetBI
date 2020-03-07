Attribute VB_Name = "m002_DataAccess"
Option Explicit


Sub GetReportList(ByRef ReportList() As TypeReportList)

    Dim i As Integer
    Dim j As Integer
    Dim lo As ListObject
    
    
    Set lo = ActiveWorkbook.Sheets("ReportList").ListObjects("tbl_ReportList")
    j = 0
    ReDim ReportList(0 To lo.DataBodyRange.Rows.Count - 1)

    With lo
        For i = 1 To .DataBodyRange.Rows.Count
            ReportList(i - 1).ReportName = .ListColumns("Report Name").DataBodyRange.Cells(i)
            ReportList(i - 1).SheetName = .ListColumns("Sheet Name").DataBodyRange.Cells(i)
            ReportList(i - 1).ReportCategory = .ListColumns("Report Category").DataBodyRange.Cells(i)
            ReportList(i - 1).RunWithRefresh = .ListColumns("Run with table refresh").DataBodyRange.Cells(i)
            ReportList(i - 1).RunWithoutRefresh = .ListColumns("Run without table refresh").DataBodyRange.Cells(i)
        Next i
    
    End With



End Sub




Sub GetReportProperties(ByVal sReportName As String, ByRef ReportProperties As TypeReportProperties)

    Dim i As Integer
    Dim lo As ListObject
    
    
    Set lo = ActiveWorkbook.Sheets("ReportProperties").ListObjects("tbl_ReportProperties")
    
    With lo
        For i = 1 To lo.DataBodyRange.Rows.Count
            If .ListColumns("Report Name").DataBodyRange.Cells(i) = sReportName Then
                ReportProperties.AutoFit = .ListColumns("AutoFit").DataBodyRange.Cells(i)
                ReportProperties.RowTotals = .ListColumns("Total Rows").DataBodyRange.Cells(i)
                ReportProperties.ColumnTotals = .ListColumns("Total Columns").DataBodyRange.Cells(i)
            End If
        Next i
    End With

End Sub



Sub GetReportFieldSettings(ByVal sReportName As String, ByRef ReportFieldSettings() As TypeReportFieldSettings)

    Dim i As Integer
    Dim lo As ListObject
    Dim j As Integer
    
    Set lo = ActiveWorkbook.Sheets("ReportFieldSettings").ListObjects("tbl_ReportFields")
    j = 0
    ReDim ReportFieldSettings(0 To MaxInt)

    With lo
        For i = 1 To lo.DataBodyRange.Rows.Count
            If .ListColumns("Report Name").DataBodyRange.Cells(i) = sReportName Then
                ReportFieldSettings(j).CubeFieldName = .ListColumns("Cube Field Name").DataBodyRange.Cells(i)
                ReportFieldSettings(j).FieldType = .ListColumns("Data Model Field Type").DataBodyRange.Cells(i)
                ReportFieldSettings(j).Orientation = .ListColumns("Orientation").DataBodyRange.Cells(i)
                ReportFieldSettings(j).Format = .ListColumns("Format").DataBodyRange.Cells(i)
                ReportFieldSettings(j).CustomFormat = .ListColumns("Custom Format").DataBodyRange.Cells(i)
                ReportFieldSettings(j).Subtotal = CBool(.ListColumns("Subtotal").DataBodyRange.Cells(i))
                ReportFieldSettings(j).SubtotalAtTop = CBool(.ListColumns("Subtotal at top").DataBodyRange.Cells(i))
                ReportFieldSettings(j).BlankLine = CBool(.ListColumns("Blank line between items").DataBodyRange.Cells(i))
                ReportFieldSettings(j).FilterType = .ListColumns("Filter Type").DataBodyRange.Cells(i)
                If ReportFieldSettings(j).FilterType <> "" Then
                    CommaSeperatedDelimit .ListColumns("Filter Values").DataBodyRange.Cells(i), ReportFieldSettings(j).FilterValues
                End If
                If .ListColumns("Collapse field values").DataBodyRange.Cells(i) <> "" Then
                    CommaSeperatedDelimit .ListColumns("Collapse field values").DataBodyRange.Cells(i), ReportFieldSettings(j).CollapseFieldValues
                End If
                j = j + 1
            End If
        Next i
    End With
    
    ReDim Preserve ReportFieldSettings(0 To j - 1)

End Sub
