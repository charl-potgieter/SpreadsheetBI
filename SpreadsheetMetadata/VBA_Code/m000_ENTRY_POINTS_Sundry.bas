Attribute VB_Name = "m000_ENTRY_POINTS_Sundry"
Option Explicit
Global Const gcsMenuName As String = "SpreadsheetBI"

Sub DisplayPopUpMenu()
Attribute DisplayPopUpMenu.VB_ProcData.VB_Invoke_Func = "M\n14"

    DeletePopUpMenu
    CreatePopUpMenu
    Application.CommandBars(gcsMenuName).ShowPopup

End Sub


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


Sub InsertStorageSheet()

    Dim Storage As ListStorage
    Dim StorageName As String
    Dim SingleCell As Range
    Dim i As Integer
    Dim Headers() As String
    Dim YesNoResponse As Integer
    
    StandardEntry
    
    If Selection.Cells.Count = 1 Then
        YesNoResponse = MsgBox("Only one cell selected, are you sure you only want one " & _
            "column in storage", vbYesNo)
    End If
    If YesNoResponse = vbNo Then GoTo Exitpoint
    
    StorageName = InputBox("Ensure range containing headers is selected (this can be later " & _
        "deleted)" & vbCrLf & _
        "Enter storage name")
    If StorageName = "" Then GoTo Exitpoint
    
    
    'Convert selected range into a 1 dimensional variant array
    Select Case True
        Case Selection.Columns.Count <> 1 And Selection.Rows.Count <> 1
            GoTo Exitpoint
        Case Selection.Rows.Count = 1
            ReDim Headers(1 To Selection.Columns.Count)
            For i = 1 To Selection.Columns.Count
                Headers(i) = Selection.Cells(i).Value
            Next i
        Case Else
            ReDim Headers(1 To Selection.Rows.Count)
            For i = 1 To Selection.Rows.Count
                Headers(i) = Selection.Cells(i).Value
            Next i
    End Select
    
    Set Storage = New ListStorage
    Storage.CreateStorage ActiveWorkbook, StorageName, Headers
    
Exitpoint:
    StandardExit

End Sub
