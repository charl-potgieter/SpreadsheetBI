Attribute VB_Name = "m000_ENTRY_POINTS_TableLoooper"
'@Folder "SpreadsheetBI"
Option Explicit
Option Private Module


Sub RunTableLooperOnActiveSheet()

    Dim ReportSheetSource As ReportingSheet
    Dim ReportSheetConsol As ReportingSheet
    Dim bReportSheetAssigned As Boolean
    
    StandardEntry
    Set ReportSheetSource = New ReportingSheet
    bReportSheetAssigned = ReportSheetSource.AssignExistingSheet(ActiveSheet)
    
    If Not bReportSheetAssigned Then
        MsgBox ("Not a valid sheet for table looping")
        GoTo Exitpoint
    End If
    
    If Not IsTableLooperSheet(ReportSheetSource.Sheet) Then
        MsgBox ("Not a valid sheet for table looping")
        GoTo Exitpoint
    End If
    
    Set ReportSheetConsol = InsertConsolLooperSheet(ReportSheetSource)
    LoopSourceAndCopyToConsolSheet ReportSheetSource, ReportSheetConsol
    FilterOutExcludedItems ReportSheetConsol
    SetLoopTableAndSheetFormat ReportSheetSource, ReportSheetConsol

    With ReportSheetConsol
        .Sheet.Range("M12").Value = "This sheet is auto produced by run looper by Spreadsheet BI AddIn"
        .Sheet.Range("M12").Font.Color = RGB(192, 0, 0)
        .Sheet.Activate
        .DefaultCursorLocation.Select
    End With

Exitpoint:
    StandardExit


End Sub
