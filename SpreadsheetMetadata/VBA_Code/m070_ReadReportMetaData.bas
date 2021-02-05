Attribute VB_Name = "m070_ReadReportMetaData"
Option Explicit
Option Private Module

Sub CreateReportMetaDataSheets()


    'Create sheet for report property capture
    If Not SheetExists(ActiveWorkbook, "ReportSheetProperties") Then
        ThisWorkbook.Sheets("ReportSheetProperties").Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
        ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count).Range("SheetHeading").Font.Color = RGB(0, 0, 0)
    End If

    
    'Create sheet for pivot table property capture
    If Not SheetExists(ActiveWorkbook, "PvtTableProperties") Then
        ThisWorkbook.Sheets("PvtTableProperties").Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
        ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count).Range("SheetHeading").Font.Color = RGB(0, 0, 0)
    End If
    
    'Create sheet for pivot field property capture
    If Not SheetExists(ActiveWorkbook, "PvtFieldProperties") Then
        ThisWorkbook.Sheets("PvtFieldProperties").Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
        ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count).Range("SheetHeading").Font.Color = RGB(0, 0, 0)
    End If
            
    
    'Set sheet order
    ActiveWorkbook.Sheets("ReportSheetProperties").Move After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    ActiveWorkbook.Sheets("PvtTableProperties").Move After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    ActiveWorkbook.Sheets("PvtFieldProperties").Move After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    
    

End Sub


Sub WritePivotTableProperties()

    Dim sht As Worksheet
    Dim lo As ListObject
    Dim rng As Range
    Dim i As Integer
    Dim pvt As PivotTable
    Dim sPivotTableProperty As String

    Set lo = ActiveWorkbook.Sheets("PvtTableProperties").ListObjects("tbl_PvtTableProperties")

    For Each sht In ActiveWorkbook.Worksheets
        If sht.PivotTables.Count = 1 Then
            Set pvt = sht.PivotTables(1)
            AddOneRowToListObject lo
            i = lo.DataBodyRange.Rows.Count
            For Each rng In lo.HeaderRowRange
                If rng = "SheetName" Then
                    lo.ListColumns(rng.Value).DataBodyRange.Cells(i) = sht.Name
                Else
                    On Error Resume Next
                    sPivotTableProperty = CallByName(pvt, rng, VbGet)
                    If Err.Number = 0 Then
                        lo.ListColumns(rng.Value).DataBodyRange.Cells(i) = sPivotTableProperty
                    Else
                        lo.ListColumns(rng.Value).DataBodyRange.Cells(i) = "VBA Error"
                    End If
                    On Error GoTo 0
                End If
            Next rng
        End If
    Next sht

End Sub







Sub temp()

    Dim s As String

    s = CallByName(ActiveSheet, "name", VbGet)
    Debug.Print (s)

End Sub


