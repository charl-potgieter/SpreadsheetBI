VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReportingSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Storage.Reporting"
Option Explicit

Private Type TypeReportingSheet
    shtReport As Worksheet
End Type

Private this As TypeReportingSheet

Private Const NumberOfHiddenWorkingRows As Integer = 4
Private Const NumberOfHiddenWorkingColumns As Integer = 4
Private Const NumberOfErrorCheckRows As Integer = 4
Private Const NumberOfErrorCheckColumns As Integer = 4


Public Property Get Heading() As String
    On Error Resume Next
    Heading = this.shtReport.Names("Heading").RefersToRange.Value
    If Err.Number <> 0 Then Heading = ""
End Property


Public Property Let Heading(ByVal sHeading As String)
    this.shtReport.Range("Heading").Value = sHeading
End Property


Public Property Get DefaultCursorLocation() As Range
    Set DefaultCursorLocation = this.shtReport.Range("DefaultCursorLocation")
End Property


Public Property Get Category() As String
    On Error Resume Next
    Category = this.shtReport.Names("Category").RefersToRange.Value
    If Err.Number <> 0 Then Category = ""
End Property


Public Property Let Category(ByVal sCategory As String)
    this.shtReport.Range("Category").Value = sCategory
End Property


Public Property Get WorkbookErrorStatusFormula() As String
    WorkbookErrorStatusFormula = this.shtReport.Names("WorkbookErrorStatus").RefersToRange.Formula
End Property


Public Property Let WorkbookErrorStatusFormula(ByVal sFormula As String)
    this.shtReport.Names("WorkbookErrorStatus").RefersToRange.Formula = sFormula
End Property


Public Property Get SheetErrorStatusFormula() As String
    SheetErrorStatusFormula = this.shtReport.Names("SheetErrorStatus").RefersToRange.Formula
End Property


Public Property Let SheetErrorStatusFormula(ByVal sFormula As String)
    this.shtReport.Names("SheetErrorStatus").RefersToRange.Formula = sFormula
End Property



Public Property Get FreezePaneRow() As Long
    
    Dim WorkbookSelectedBeforeSub As Workbook
    Dim SelectedSheetOnReportingWorkbookBeforeSub As Worksheet
    
    Set WorkbookSelectedBeforeSub = ActiveWorkbook
    this.shtReport.Parent.Activate
    Set SelectedSheetOnReportingWorkbookBeforeSub = ActiveSheet
    this.shtReport.Activate
    
    FreezePaneRow = ActiveWindow.SplitRow
    
    SelectedSheetOnReportingWorkbookBeforeSub.Activate
    WorkbookSelectedBeforeSub.Activate
    
End Property


Public Property Let FreezePaneRow(ByVal lFreezeRow As Long)
    
    Dim WorkbookSelectedBeforeSub As Workbook
    Dim SelectedSheetOnReportingWorkbookBeforeSub As Worksheet
    Dim lFreezeCol As Long
    Dim wkbToFreeze As Workbook
    
    Set WorkbookSelectedBeforeSub = ActiveWorkbook
    Set wkbToFreeze = this.shtReport.Parent
    wkbToFreeze.Activate
    Set SelectedSheetOnReportingWorkbookBeforeSub = ActiveSheet
    this.shtReport.Activate
    
    With ActiveWindow
        lFreezeCol = ActiveWindow.SplitColumn
        .FreezePanes = False
        .Split = False
        FreezeActiveWindowPanes lFreezeRow, lFreezeCol
    End With
    
    SelectedSheetOnReportingWorkbookBeforeSub.Activate
    WorkbookSelectedBeforeSub.Activate
    
End Property


Public Property Get FreezePaneCol() As Long
    
    Dim WorkbookSelectedBeforeSub As Workbook
    Dim SelectedSheetOnReportingWorkbookBeforeSub As Worksheet
    
    Set WorkbookSelectedBeforeSub = ActiveWorkbook
    this.shtReport.Parent.Activate
    Set SelectedSheetOnReportingWorkbookBeforeSub = ActiveSheet
    this.shtReport.Activate
    FreezePaneCol = ActiveWindow.SplitColumn
    
    SelectedSheetOnReportingWorkbookBeforeSub.Activate
    WorkbookSelectedBeforeSub.Activate
    
End Property


Public Property Let FreezePaneCol(ByVal lFreezeColumn As Long)
    
    Dim WorkbookSelectedBeforeSub As Workbook
    Dim SelectedSheetOnReportingWorkbookBeforeSub As Worksheet
    Dim lFreezeRow As Long
    
    Set WorkbookSelectedBeforeSub = ActiveWorkbook
    this.shtReport.Parent.Activate
    Set SelectedSheetOnReportingWorkbookBeforeSub = ActiveSheet
    this.shtReport.Activate
    
    With ActiveWindow
        lFreezeRow = ActiveWindow.SplitRow
        .FreezePanes = False
        .Split = False
        FreezeActiveWindowPanes lFreezeRow, lFreezeColumn
    End With
    
    SelectedSheetOnReportingWorkbookBeforeSub.Activate
    WorkbookSelectedBeforeSub.Activate
    
End Property


Public Property Get Sheet() As Worksheet
    Set Sheet = this.shtReport
End Property


Public Property Get Name() As String
    Name = this.shtReport.Name
End Property


Public Property Get HiddenRangesAreVisible() As Boolean
    HiddenRangesAreVisible = this.shtReport.Range("HiddenWorkingRows").EntireRow.Hidden = False
End Property


Public Property Let Name(ByVal sName As String)

    Dim Counter As Long
    Dim OverMaxSheetLength As Integer
    Dim TryName As String
    Counter = 0
    
    'Get first availabe sheet name add numeric prefix if name is taken
    TryName = sName
    On Error Resume Next
    Do While this.shtReport.Name <> TryName
        If Counter <> 0 Then
            TryName = sName & "_" & Counter
            If Len(TryName) > 31 Then
                OverMaxSheetLength = 31 - Len(TryName)
                TryName = Left(sName, OverMaxSheetLength) & "_" & Counter
            End If
        End If
        this.shtReport.Name = TryName
         Counter = Counter + 1
    Loop
    On Error GoTo 0
        
End Property


Public Property Let SheetFont(ByVal FontName As String)
    this.shtReport.Cells.Font.Name = FontName
End Property


Public Property Let DefaultFontSize(ByVal DefaultSize As String)
    this.shtReport.Cells.Font.Size = DefaultSize
End Property


Public Property Let ZoomPercentage(ByVal Pct As Double)
    this.shtReport.Activate
    ActiveWindow.Zoom = Pct
End Property


Public Property Let HeadingFontColour(ByVal RgbArray)
    this.shtReport.Range("Heading").Font.Color = RGB(RgbArray(0), RgbArray(1), RgbArray(2))
End Property


Public Property Let HeadingFontSize(ByVal FontSize As Integer)
    this.shtReport.Range("Heading").Font.Size = FontSize
End Property



Public Sub Create(ByVal wkb As Workbook, _
    Optional ByVal lSheetPosition As Long = 0)
    
    InsertReportingSheet wkb, lSheetPosition
    SetRangeNames
    FormatSheet
    Me.Category = "Category"
    Me.Heading = "Heading - " & this.shtReport.Name
    
End Sub


Public Function AssignExistingSheet(ByVal sht As Worksheet) As Boolean
'Only assigns sheet if it already meets reporting sheet criteria
    
    Dim bValidSheet As Boolean
    Dim test As Variant
    
    On Error Resume Next
    test = sht.Range("HiddenWorkingRows").Address
    test = sht.Range("HiddenWorkingColumns").Address
    test = sht.Range("ErrorCheckColumns").Address
    test = sht.Range("ErrorCheckRows").Address
    test = sht.Range("WorkbookErrorStatus").Address
    test = sht.Range("Category").Address
    test = sht.Range("Heading").Address
    test = sht.Range("ReturnToIndex").Address
    test = sht.Range("DefaultCursorLocation").Address
    bValidSheet = (Err.Number = 0)
    On Error GoTo 0
    
    If bValidSheet Then Set this.shtReport = sht
    AssignExistingSheet = bValidSheet

End Function



Public Function CreateFromExistingSheet(ByVal sht As Worksheet) As Boolean
'Converts sht to a reporting sheet
    
    CreateFromExistingSheet = (this.shtReport Is Nothing)

    If CreateFromExistingSheet Then
        Range(Columns(1), _
            Columns(NumberOfHiddenWorkingColumns + NumberOfErrorCheckColumns + 1)). _
            Insert xlShiftToRight
        Range(Rows(1), _
            Rows(NumberOfHiddenWorkingRows + NumberOfErrorCheckRows + 5)). _
            Insert xlShiftDown
        Set this.shtReport = sht
        SetRangeNames
        FormatSheet
        If Me.Category = "" Then Me.Category = "Category"
        If Me.Heading = "" Then Me.Heading = "Heading - " & this.shtReport.Name
    End If

End Function


Public Sub Delete()
    this.shtReport.Delete
End Sub


Private Sub InsertReportingSheet(ByVal wkb As Workbook, _
    Optional ByVal lSheetPosition As Long = 0)
    
    Dim NumberOfSheetsInWorkbook As Long
    
    If lSheetPosition <> 0 Then
        Set this.shtReport = wkb.Sheets.Add _
            (After:=wkb.Worksheets.Item(lSheetPosition))
    Else
        NumberOfSheetsInWorkbook = wkb.Sheets.Count
        Set this.shtReport = wkb.Sheets.Add _
            (After:=wkb.Worksheets.Item(NumberOfSheetsInWorkbook))
    End If
        
End Sub


Private Sub SetRangeNames()
    
    With this.shtReport
        
        On Error Resume Next
        .Names("HiddenWorkingRows").Delete
        .Names("HiddenWorkingColumns").Delete
        .Names("ErrorCheckRows").Delete
        .Names("ErrorCheckColumns").Delete
        .Names("Category").Delete
        .Names("Heading").Delete
        .Names("WorkbookErrorStatus").Delete
        .Names("ReturnToIndex").Delete
        .Names("DefaultCursorLocation").Delete
        On Error GoTo 0
        
        .Names.Add Name:="HiddenWorkingRows", RefersTo:="=$1:$" & NumberOfHiddenWorkingRows
        
        .Names.Add Name:="HiddenWorkingColumns", RefersTo:="=" & _
            .Range(Columns(1), Columns(NumberOfHiddenWorkingColumns)).Address
            
        .Names.Add Name:="ErrorCheckRows", RefersTo:="=" & _
            Range(NumberOfHiddenWorkingRows + 1 & ":" & _
            NumberOfHiddenWorkingRows + NumberOfErrorCheckRows).Address
        
        .Names.Add Name:="ErrorCheckColumns", RefersTo:="=" & _
            Range(Columns(NumberOfHiddenWorkingColumns + 1), _
            Columns(NumberOfHiddenWorkingColumns + NumberOfErrorCheckColumns)).Address
            
        .Names.Add Name:="SheetErrorStatus", RefersTo:="=$C$1"
        .Names.Add Name:="WorkbookErrorStatus", RefersTo:="=$C$2"
        .Names.Add Name:="Category", RefersTo:="=$C$3"
        
        
        .Names.Add Name:="Heading", RefersTo:="=" & _
            .Cells(1, 1).Offset(NumberOfErrorCheckRows + NumberOfErrorCheckRows + 1, _
            NumberOfHiddenWorkingRows + NumberOfHiddenWorkingColumns + 1).Address
            
        .Names.Add Name:="ReturnToIndex", RefersTo:="=" & _
            .Range("Heading").Offset(1, 0).Address
            
        .Names.Add Name:="DefaultCursorLocation", RefersTo:="=" & _
            .Range("Heading").Offset(-1, -1).Address
            
        .Names.Add Name:="Notes", RefersTo:="=" & _
            .Range(.Cells(1, 1), .Cells((NumberOfHiddenWorkingRows + NumberOfErrorCheckRows), _
            (NumberOfHiddenWorkingColumns + NumberOfErrorCheckRows))).Address
            
    End With
    
End Sub


Public Sub ToggleErrorCheckRangeVisbility(ByVal ShowHiddenRange As Boolean)

    With this.shtReport
        .Activate
        If ShowHiddenRange Then
            .Range("HiddenWorkingRows").EntireRow.Hidden = False
            .Range("HiddenWorkingColumns").EntireColumn.Hidden = False
            .Range("ErrorCheckRows").EntireRow.Hidden = False
            .Range("ErrorCheckColumns").EntireColumn.Hidden = False
            .Range("A1").Select
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        Else
            .Range("HiddenWorkingRows").EntireRow.Hidden = True
            .Range("HiddenWorkingColumns").EntireColumn.Hidden = True
            .Range("ErrorCheckRows").EntireRow.Hidden = True
            .Range("ErrorCheckColumns").EntireColumn.Hidden = True
            .Range("DefaultCursorLocation").Select
        End If
        
    End With
    
End Sub






Private Sub FormatSheet()

    Dim wkbReportWorkbook As Workbook
    Dim wkbActivePriorToSub As Workbook

    'Necessary to activate wkbReportWorkbook to set zoom and no gridlines
    Set wkbActivePriorToSub = ActiveWorkbook
    Set wkbReportWorkbook = this.shtReport.Parent
    wkbReportWorkbook.Activate
    this.shtReport.Select

    With this.shtReport
        .DisplayPageBreaks = False
        .Columns(NumberOfHiddenWorkingColumns + NumberOfErrorCheckColumns + 1).ColumnWidth = 4

        .Range("Heading").Font.Bold = True

        .Range("Category").Font.Color = RGB(170, 170, 170)
        .Range("Category").Font.Size = 8

        .Range("WorkbookErrorStatus").Font.Color = RGB(170, 170, 170)
        .Range("SheetErrorStatus").Font.Color = RGB(170, 170, 170)
        .Range("ErrorCheckRows").Cells(1).Font.Bold = True

        .Range("HiddenWorkingRows").Interior.Color = RGB(217, 217, 217)
        .Range("HiddenWorkingColumns").Interior.Color = RGB(217, 217, 217)
        .Range("ErrorCheckRows").Interior.Color = RGB(117, 117, 117)
        .Range("ErrorCheckColumns").Interior.Color = RGB(117, 117, 117)
        .Range("ErrorCheckRows").Font.Color = RGB(255, 255, 255)
        .Range("ErrorCheckColumns").Font.Color = RGB(255, 255, 255)
        .Range("ErrorCheckRows").Font.Bold = True
        .Range("ErrorCheckColumns").Font.Bold = True
            
        .Range("Notes").Interior.Color = RGB(205, 225, 243)
        .Range("Notes").Font.Color = RGB(0, 0, 0)
        .Range("Notes").EntireColumn.ColumnWidth = 10
        .Range("Notes").BorderAround Weight:=xlThick
        .Range("Notes").Cells(1, 1).Value = "Sheet error status"
        .Range("Notes").Cells(2, 1).Value = "Workbook error status"
        .Range("Notes").Cells(3, 1).Value = "Category"
        .Range("Notes").Cells(6, 1).Value = "Hidden workings in light grey"
        .Range("Notes").Cells(7, 1).Value = "Hidden error checks in dark grey"
        .Range("Notes").Cells(2, 2).Font.Bold = True
        .Range("Notes").Cells(3, 2).Font.Bold = True

        .Range("ErrorCheckRows").EntireRow.Hidden = True
        .Range("ErrorCheckColumns").EntireColumn.Hidden = True
        .Range("HiddenWorkingRows").EntireRow.Hidden = True
        .Range("HiddenWorkingColumns").EntireColumn.Hidden = True
        
        .Range("Category").Font.Color = RGB(0, 0, 255)
        .Range("Category").Font.Bold = True
        
        SetErrorConditionalFormatting
        ActiveWindow.DisplayGridlines = False
                
    End With

    'Reactivate the original active workbook
    wkbActivePriorToSub.Activate
    
    Set wkbReportWorkbook = Nothing
    Set wkbActivePriorToSub = Nothing
    
End Sub



Private Sub SetErrorConditionalFormatting()

    Dim HeadingCell As Range
    Dim FormatConditionSheetError
    Dim FormatConditionWorkbookError As FormatCondition
    
    Set HeadingCell = this.shtReport.Range("Heading")
    HeadingCell.FormatConditions.Delete
    
    Set FormatConditionSheetError = HeadingCell.FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=SheetErrorStatus<>""OK""")
    
    With FormatConditionSheetError
        .NumberFormat = """SHEET ERROR""; ""SHEET ERROR"";" & _
            """SHEET ERROR""; ""SHEET ERROR"""
        .Font.Bold = True
        .Font.Italic = False
        .Font.Color = RGB(255, 0, 0)
        .Font.TintAndShade = 0
    End With
    
    Set FormatConditionWorkbookError = HeadingCell.FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=WorkbookErrorStatus<>""OK""")
    
    With FormatConditionWorkbookError
        .NumberFormat = """WORKBOOK ERROR""; ""WORKBOOK ERROR"";" & _
            """WORKBOOK ERROR""; ""WORKBOOK ERROR"""
        .Font.Bold = True
        .Font.Italic = False
        .Font.Color = RGB(255, 0, 0)
        .Font.TintAndShade = 0
    End With
    
    
End Sub



Private Sub FreezeActiveWindowPanes(ByVal lFreezeRow As Long, ByVal lFreezeColumn As Long)
   
    Const MinScreenPercent As Double = 0.25
   
    With ActiveWindow
    
        'Freeze panes does not seem to work if window is too small
        If .Width < ScreenDimensionWidth * MinScreenPercent Then
            .Width = ScreenDimensionWidth * MinScreenPercent
        End If
        If .Height < ScreenDimensionHeight * MinScreenPercent Then
            .Height = ScreenDimensionHeight * MinScreenPercent
        End If
        
        Select Case True
        Case lFreezeColumn = 0 And lFreezeRow <> 0
            .SplitRow = lFreezeRow
            .FreezePanes = True
        Case lFreezeColumn <> 0 And lFreezeRow = 0
            .SplitColumn = lFreezeColumn
            .FreezePanes = True
        Case lFreezeColumn <> 0 And lFreezeRow <> 0
            .SplitColumn = lFreezeColumn
            .SplitRow = lFreezeRow
            .FreezePanes = True
        End Select
    End With

End Sub





Private Sub Class_Terminate()
    Set this.shtReport = Nothing
End Sub
