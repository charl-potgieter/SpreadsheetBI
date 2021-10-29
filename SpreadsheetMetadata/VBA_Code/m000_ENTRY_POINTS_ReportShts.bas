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
    ReportSht.SheetFont = ReportSheetFormat.Item("Sheet font")
    ReportSht.DefaultFontSize = ReportSheetFormat.Item("Default font size")
    ReportSht.ZoomPercentage = ReportSheetFormat.Item("Zoom percentage")
    ReportSht.HeadingFontColour = Array( _
        ReportSheetFormat.Item("Heading colour red (0 to 255)"), _
        ReportSheetFormat.Item("Heading colour green (0 to 255)"), _
        ReportSheetFormat.Item("Heading colour blue (0 to 255)"))
    ReportSht.HeadingFontSize = ReportSheetFormat.Item("Heading font size")
    
    InsertIndexPage ActiveWorkbook
    ReportSht.Sheet.Activate
    ReportSht.DefaultCursorLocation.Select

    Set wkb = Nothing
    Set ReportSht = Nothing
    StandardExit

End Function


Sub ConvertSelectedSheetsToReportingSheet()

    Dim ReportSht As ReportingSheet
    Dim ReportSheetFormat As Dictionary
    Dim SelectedSheetNames() As String
    Dim i As Integer

    StandardEntry
    Set ReportSheetFormat = GetSavedReportSheetFormat
    
    ReDim SelectedSheetNames(1 To ActiveWindow.SelectedSheets.Count)
    For i = 1 To ActiveWindow.SelectedSheets.Count
        SelectedSheetNames(i) = ActiveWindow.SelectedSheets(i).Name
    Next i
    
    For i = LBound(SelectedSheetNames) To UBound(SelectedSheetNames)
    
        'Difficult to achieve without selecting as sheet needs to be active to set zoom
        ActiveWorkbook.Sheets(SelectedSheetNames(i)).Select
        
        Set ReportSht = New ReportingSheet
        ReportSht.CreateFromExistingSheet ActiveSheet
        ReportSht.SheetFont = ReportSheetFormat.Item("Sheet font")
        ReportSht.DefaultFontSize = ReportSheetFormat.Item("Default font size")
        ReportSht.ZoomPercentage = ReportSheetFormat.Item("Zoom percentage")
        ReportSht.HeadingFontColour = Array( _
            ReportSheetFormat.Item("Heading colour red (0 to 255)"), _
            ReportSheetFormat.Item("Heading colour green (0 to 255)"), _
            ReportSheetFormat.Item("Heading colour blue (0 to 255)"))
        ReportSht.HeadingFontSize = ReportSheetFormat.Item("Heading font size")
    Next i
    
    InsertIndexPage ActiveWorkbook
    ReportSht.Sheet.Activate
    
    Set ReportSht = Nothing
    Set ReportSheetFormat = Nothing
    
    StandardExit

End Sub




Sub ToggleErrorCheckRangeVisbilityOnSelectedSheets()
Attribute ToggleErrorCheckRangeVisbilityOnSelectedSheets.VB_ProcData.VB_Invoke_Func = "H\n14"

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
Attribute IndexPageNavigation.VB_ProcData.VB_Invoke_Func = "I\n14"

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

    Dim uf As uf_ReportSheetFormat
    Dim ReportSheetFormatDict(0 To 6) As Dictionary
    Dim i As Integer
    
    StandardEntry
    Set uf = New uf_ReportSheetFormat
    
    uf.tbSheetFont.Value = GetReportSheetFormatItem("Sheet Font")
    uf.tbDefaultFontSize.Value = GetReportSheetFormatItem("Default font size")
    uf.tbZoomPercentage.Value = GetReportSheetFormatItem("Zoom percentage")
    uf.tbHeadingColourRed.Value = GetReportSheetFormatItem("Heading colour red (0 to 255)")
    uf.tbHeadingColourGreen.Value = GetReportSheetFormatItem("Heading colour green (0 to 255)")
    uf.tbHeadingColourBlue.Value = GetReportSheetFormatItem("Heading colour blue (0 to 255)")
    uf.tbHeadingFontSize.Value = GetReportSheetFormatItem("Heading font size")
    
    uf.Show
    If uf.UserCancelled Then GoTo Exitpoint
    
    For i = LBound(ReportSheetFormatDict) To UBound(ReportSheetFormatDict)
        Set ReportSheetFormatDict(i) = New Dictionary
    Next i
    
    ReportSheetFormatDict(0).Add "Item", "Sheet font"
    ReportSheetFormatDict(0).Add "Value", uf.tbSheetFont.Value

    ReportSheetFormatDict(1).Add "Item", "Default font size"
    ReportSheetFormatDict(1).Add "Value", uf.tbDefaultFontSize.Value
    
    ReportSheetFormatDict(2).Add "Item", "Zoom percentage"
    ReportSheetFormatDict(2).Add "Value", uf.tbZoomPercentage.Value

    ReportSheetFormatDict(3).Add "Item", "Heading colour red (0 to 255)"
    ReportSheetFormatDict(3).Add "Value", uf.tbHeadingColourRed.Value

    ReportSheetFormatDict(4).Add "Item", "Heading colour green (0 to 255)"
    ReportSheetFormatDict(4).Add "Value", uf.tbHeadingColourGreen.Value

    ReportSheetFormatDict(5).Add "Item", "Heading colour blue (0 to 255)"
    ReportSheetFormatDict(5).Add "Value", uf.tbHeadingColourBlue.Value
    
    ReportSheetFormatDict(6).Add "Item", "Heading font size"
    ReportSheetFormatDict(6).Add "Value", uf.tbHeadingFontSize

    WriteReportSheetFormat ReportSheetFormatDict
    ThisWorkbook.Save

Exitpoint:
    Unload uf
    Set uf = Nothing
    For i = LBound(ReportSheetFormatDict) To UBound(ReportSheetFormatDict)
        Set ReportSheetFormatDict(i) = Nothing
    Next i
    StandardExit

End Sub



