Attribute VB_Name = "m000_ENTRY_POINTS_ReportShts"
Option Explicit


Sub InsertIndexPageActiveWorkbook()
    
    Dim IndexSheet As Worksheet

    StandardEntry
    Set IndexSheet = InsertIndexPage(ActiveWorkbook)
    IndexSheet.Activate
    IndexSheet.Range("DefaultCursorLocation").Select
    Set IndexSheet = Nothing
    StandardExit

End Sub


Function InsertReportingSheetSheetIntoActiveWorkbook()
    
    Dim ReportSht As ReportingSheet
    Dim wkb As Workbook
    Dim ReportSheetFormat As Dictionary

    StandardEntry

    Set wkb = ActiveWorkbook
    Set ReportSht = New ReportingSheet
    ReportSht.Create wkb, ActiveSheet.Index
    
    Set ReportSheetFormat = GetSavedReportSheetFormat
    ReportSht.SheetFont = ReportSheetFormat.item("Sheet font")
    ReportSht.DefaultFontSize = ReportSheetFormat.item("Default font size")
    ReportSht.ZoomPercentage = ReportSheetFormat.item("Zoom percentage")
    ReportSht.HeadingFontColour = Array( _
        ReportSheetFormat.item("Heading colour red (0 to 255)"), _
        ReportSheetFormat.item("Heading colour green (0 to 255)"), _
        ReportSheetFormat.item("Heading colour blue (0 to 255)"))
    ReportSht.HeadingFontSize = ReportSheetFormat.item("Heading font size")
    
    InsertIndexPage ActiveWorkbook
    ReportSht.Sheet.Activate
    ReportSht.DefaultCursorLocation.Select

    Set wkb = Nothing
    Set ReportSht = Nothing
    StandardExit

End Function


Sub ConvertActiveSheetToReportingSheet()

    Dim ReportSht As ReportingSheet
    Dim ReportSheetFormat As Dictionary

    StandardEntry
    Set ReportSht = New ReportingSheet

    ReportSht.CreateFromExistingSheet ActiveSheet
    Set ReportSheetFormat = GetSavedReportSheetFormat
    ReportSht.SheetFont = ReportSheetFormat.item("Sheet font")
    ReportSht.DefaultFontSize = ReportSheetFormat.item("Default font size")
    ReportSht.ZoomPercentage = ReportSheetFormat.item("Zoom percentage")
    ReportSht.HeadingFontColour = Array( _
        ReportSheetFormat.item("Heading colour red (0 to 255)"), _
        ReportSheetFormat.item("Heading colour green (0 to 255)"), _
        ReportSheetFormat.item("Heading colour blue (0 to 255)"))
    ReportSht.HeadingFontSize = ReportSheetFormat.item("Heading font size")
    
    InsertIndexPage ActiveWorkbook
    ReportSht.Sheet.Activate
    StandardExit

End Sub




Sub ToggleErrorCheckRangeVisbilityOnSelectedSheets()

    Dim sht As Worksheet
    Dim ReportSheet As ReportingSheet
    Dim ReportIsAssigned As Boolean
    Dim obj As Object
    Dim ShowHiddenRange As Boolean
    Dim CurrentlyActiveSheet As Worksheet
    Dim IsFirstReportingSheetInSelection As Boolean
    
    StandardEntry
    Set CurrentlyActiveSheet = ActiveSheet
    IsFirstReportingSheetInSelection = True
    
    'Toggling can occur for multiple selected sheets
    'Visibility is set based on status of the first sheet
    For Each obj In ActiveWindow.SelectedSheets
        Set sht = obj
        Set ReportSheet = New ReportingSheet
        ReportIsAssigned = ReportSheet.AssignExistingSheet(sht)
        If ReportIsAssigned Then
            If IsFirstReportingSheetInSelection Then
                ShowHiddenRange = Not ReportSheet.HiddenRangesAreVisible
                IsFirstReportingSheetInSelection = False
            End If
            ReportSheet.ToggleErrorCheckRangeVisbility ShowHiddenRange
        End If
        Set ReportSheet = Nothing
    Next obj
    Set obj = Nothing
    CurrentlyActiveSheet.Activate
    StandardExit
    
End Sub



Sub IndexPageNavigation()

    Dim wkb As Workbook
    Dim TargetSheetName As String
    Dim TargetSheet As Worksheet

    Set wkb = ActiveWorkbook

    Select Case True
    
        Case ActiveSheet.Name <> "Index" And SheetExists(wkb, "Index")
            Sheets("Index").Activate
            On Error Resume Next
            On Error GoTo 0
            
        Case Selection.Rows.Count <> 1
            'Do Nothing
        
        Case wkb.Sheets("Index").Range("HiddenSheetNamesCol").Cells(Selection.Row) = ""
            'Do Nothing
            
        Case Else
            On Error Resume Next
            TargetSheetName = wkb.Sheets("Index").Range("HiddenSheetNamesCol").Cells(Selection.Row)
            Set TargetSheet = wkb.Sheets(TargetSheetName)
            TargetSheet.Activate
            On Error GoTo 0
            
    End Select

End Sub



Sub SetReportSheetFormat()

    Dim ReportSheetFormatStorage As ListStorage
    Dim wkbUserInput As Workbook
    Dim UserInputSheet As Worksheet
    Dim UserInputListObj As ListObject

    StandardEntry
    Set ReportSheetFormatStorage = New ListStorage
    ReportSheetFormatStorage.AssignStorage ThisWorkbook, "ReportSheetFormat"

    Set wkbUserInput = Application.Workbooks.Add
    Set UserInputSheet = wkbUserInput.Sheets(1)
    FormatSheet UserInputSheet
    UserInputSheet.Range("SheetHeading") = "Report Sheet Formating --> Run 'Format->" & _
        "Save default report sheet format"
    UserInputSheet.Range("SheetCategory") = ""
    
    ReportSheetFormatStorage.ListObj.Range.Copy
    UserInputSheet.Activate
    UserInputSheet.Range("B5").PasteSpecial xlPasteValues
    
    Set UserInputListObj = UserInputSheet.ListObjects.Add(xlSrcRange, Range("$B$5").CurrentRegion, , xlYes)
    FormatTable UserInputListObj
    UserInputListObj.Name = "tbl_ReportSheetFormat"

    ActiveWindow.WindowState = xlMaximized

ExitPoint:
    Set ReportSheetFormatStorage = Nothing
    StandardExit

End Sub


Sub SaveReportSheetFormat()

    Dim i As Integer
    Dim ReportSheetFormatStorage As ListStorage
    Dim UserInputListObj As ListObject

    StandardEntry
    Set ReportSheetFormatStorage = New ListStorage
    ReportSheetFormatStorage.AssignStorage ThisWorkbook, "ReportSheetFormat"
    Set UserInputListObj = ActiveWorkbook.Worksheets(1).ListObjects("tbl_ReportSheetFormat")
    
    For i = 1 To ReportSheetFormatStorage.NumberOfRecords
        ReportSheetFormatStorage.ListObj.ListColumns("Value").DataBodyRange.Cells(i) = _
            UserInputListObj.ListColumns("Value").DataBodyRange.Cells(i)
    Next i

    MsgBox ("Report sheet format updated")
    
    UserInputListObj.Parent.Parent.Close
    
    ThisWorkbook.Save
    StandardExit

End Sub


