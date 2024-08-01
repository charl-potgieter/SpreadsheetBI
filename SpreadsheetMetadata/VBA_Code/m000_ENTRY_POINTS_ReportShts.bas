Attribute VB_Name = "m000_ENTRY_POINTS_ReportShts"
Option Explicit


Sub InsertIndexPageActiveWorkbook()

    Dim WorkbookIndex As IndexPage

    StandardEntry
    Set WorkbookIndex = New IndexPage
    
    WorkbookIndex.Create ActiveWorkbook
    WorkbookIndex.ActivateAtDefauiltCursorLocation
    Set WorkbookIndex = Nothing
    StandardExit

End Sub


Function InsertReportingSheetSheetIntoActiveWorkbook()
    
    Dim ReportSht As ReportingSheet
    Dim CurrentReportSheet As ReportingSheet
    Dim WorkbookIndex As IndexPage
    Dim CurrentSheetAssigned As Boolean
    Dim Category As String
    Dim Header As String
    Dim SheetName As String
    Dim wkb As Workbook
    
    StandardEntry

    Set wkb = ActiveWorkbook
    
    'Get header of current sheet if  it is a report sheet
    Set CurrentReportSheet = New ReportingSheet
    CurrentSheetAssigned = CurrentReportSheet.AssignExistingSheet(ActiveSheet)
    If CurrentSheetAssigned Then
        Category = CurrentReportSheet.Category
    End If
    
    Set ReportSht = New ReportingSheet
    ReportSht.Create wkb, ActiveSheet.Index
    
    Category = InputBox(Prompt:="Enter sheet category", Default:=Category)
    ReportSht.Category = Category
    
    Header = InputBox(Prompt:="Enter sheet heading")
    ReportSht.Heading = Header
    
    SheetName = InputBox(Prompt:="Enter sheet name", _
        Default:=Replace(WorksheetFunction.Proper(Header), " ", ""))
    ReportSht.Name = SheetName
    
    Set WorkbookIndex = New IndexPage
    WorkbookIndex.Create wkb
    ReportSht.Sheet.Activate
    ReportSht.DefaultCursorLocation.Select

    Set WorkbookIndex = Nothing
    Set wkb = Nothing
    Set ReportSht = Nothing
    StandardExit

End Function


Sub ConvertSelectedSheetsToReportingSheet()

    Dim ReportSht As ReportingSheet
    Dim WorkbookIndex As IndexPage
    Dim ReportSheetformat As Dictionary
    Dim SelectedSheetNames() As String
    Dim i As Integer

    StandardEntry
    Set ReportSheetformat = GetSavedReportSheetFormat
    
    ReDim SelectedSheetNames(1 To ActiveWindow.SelectedSheets.Count)
    For i = 1 To ActiveWindow.SelectedSheets.Count
        SelectedSheetNames(i) = ActiveWindow.SelectedSheets(i).Name
    Next i
    
    For i = LBound(SelectedSheetNames) To UBound(SelectedSheetNames)
    
        'Difficult to achieve without selecting as sheet needs to be active to set zoom
        ActiveWorkbook.Sheets(SelectedSheetNames(i)).Select
        
        Set ReportSht = New ReportingSheet
        ReportSht.CreateFromExistingSheet ActiveSheet
        ReportSht.SheetFont = ReportSheetformat.item("Sheet font")
        ReportSht.DefaultFontSize = ReportSheetformat.item("Default font size")
        ReportSht.ZoomPercentage = ReportSheetformat.item("Zoom percentage")
        ReportSht.HeadingFontColour = Array( _
            ReportSheetformat.item("Heading colour red (0 to 255)"), _
            ReportSheetformat.item("Heading colour green (0 to 255)"), _
            ReportSheetformat.item("Heading colour blue (0 to 255)"))
        ReportSht.HeadingFontSize = ReportSheetformat.item("Heading font size")
    Next i
    
    Set WorkbookIndex = New IndexPage
    WorkbookIndex.Create ActiveWorkbook
    ReportSht.Sheet.Activate
    
    Set WorkbookIndex = Nothing
    Set ReportSht = Nothing
    Set ReportSheetformat = Nothing
    
    StandardExit

End Sub





Sub IndexPageNavigation()
Attribute IndexPageNavigation.VB_ProcData.VB_Invoke_Func = "I\n14"

    Dim wkb As Workbook
    Dim TargetSheetName As String
    Dim TargetSheet As Worksheet
    Dim ReportSheet As ReportingSheet
    Dim ReportSheetAssigned As Boolean
    

    Set wkb = ActiveWorkbook
    Set ReportSheet = New ReportingSheet
    ReportSheetAssigned = ReportSheet.AssignExistingSheet(ActiveSheet)


    Select Case True
    
        Case ReportSheetAssigned
            ReportSheet.Sheet.Range("ReturnToIndex").Hyperlinks(1).Follow
    
        Case ActiveSheet.Name <> "Index" And SheetExists(wkb, "Index")
            Sheets("Index").Activate
            
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



