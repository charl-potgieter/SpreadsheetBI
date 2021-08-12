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

Private Type TReportingSheet
    shtReport As Worksheet
End Type

Private this As TReportingSheet
Private Const csHeadingAddress As String = "=$F$10"
Private Const csCategoryAddress As String = "=$E$9"
Private Const csErrorCheckRows As String = "=$A:$D"
Private Const csErrorCheckColumns As String = "=$1:$8"
Private Const csWorkbookErrorStatus As String = "=$I$9"


Public Property Get Heading() As String
    On Error Resume Next
    Heading = this.shtReport.Names("Heading").RefersToRange.Value
    If Err.Number <> 0 Then Heading = ""
End Property


Public Property Let Heading(ByVal sHeading As String)
    this.shtReport.Range("Heading").Value = sHeading
End Property


Public Property Get Category() As String
    On Error Resume Next
    Category = this.shtReport.Names("Category").RefersToRange.Value
    If Err.Number <> 0 Then Category = ""
End Property


Public Property Let Category(ByVal sCategory As String)
    this.shtReport.Range("Category").Value = sCategory
End Property


Public Property Get WorkbooksErrorStatus() As String
    On Error Resume Next
    WorkbookErrorStatus = this.shtReport.Names("WorkbookErrorStatus").RefersToRange.Value
    If Err.Number <> 0 Then Category = ""
End Property


Public Property Let WorkbookErrorStatus(ByVal sErrorStatus As String)
    this.shtReport.Range("WorkbookErrorStatus").Value = sErrorStatus
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


Public Sub Create(ByVal wkb As Workbook, _
    Optional ByVal lSheetPosition As Long = 0)
    
    InsertReportingSheet wkb, lSheetPosition
    SetReportingSheetRangeNames
    FormatReportingSheet
    Me.Category = "Category"
    Me.Heading = "Heading"
    this.shtReport.Range("WorkbookErrorStatus").Formula = WorkbookErrorCheckFormula
    
End Sub


Public Function AssignExistingSheet(ByVal sht As Worksheet) As Boolean
'Only assigns sheet if it already meets reporting sheet criteria
    
    Dim bValidSheet As Boolean
    Dim test As Variant
    
    On Error Resume Next
    test = sht.Range("Category").Address
    test = sht.Range("Heading").Address
    test = sht.Range("WorkbookErrorStatus").Address
    test = sht.Range("ErrorCheckColumns").Address
    test = sht.Range("ErrorCheckRows").Address
    bValidSheet = (Err.Number = 0)
    On Error GoTo 0
    
    If bValidSheet Then Set this.shtReport = sht
    AssignExistingSheet = bValidSheet

End Function



Public Function CreateFromExistingSheet(ByVal sht As Worksheet) As Boolean
'Converts sht to a reporting sheet
    
    CreateFromExistingSheet = (this.shtReport Is Nothing)

    If CreateFromExistingSheet Then
        Set this.shtReport = sht
        SetReportingSheetRangeNames
        FormatReportingSheet
        If Me.Category = "" Then Me.Category = "Category"
        If Me.Heading = "" Then Me.Heading = "Heading"
        sht.Range("WorkbookErrorStatus").Formula = WorkbookErrorCheckFormula
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
            (After:=wkb.Worksheets.item(lSheetPosition))
    Else
        NumberOfSheetsInWorkbook = wkb.Sheets.Count
        Set this.shtReport = wkb.Sheets.Add _
            (After:=wkb.Worksheets.item(NumberOfSheetsInWorkbook))
    End If
        
End Sub


Private Sub SetReportingSheetRangeNames()
    
    With this.shtReport
        On Error Resume Next
        .Names("Heading").Delete
        .Names("Category").Delete
        .Names("WorkbookErrorStatus").Delete
        On Error GoTo 0
        .Names.Add Name:="Heading", RefersTo:=csHeadingAddress
        .Names.Add Name:="Category", RefersTo:=csCategoryAddress
        .Names.Add Name:="ErrorCheckColumns", RefersTo:=csErrorCheckColumns
        .Names.Add Name:="ErrorCheckRows", RefersTo:=csErrorCheckRows
        .Names.Add Name:="WorkbookErrorStatus", RefersTo:=csWorkbookErrorStatus
    End With
    
End Sub


Public Sub ToggleErrorCheckRangeVisbility()

    With this.shtReport

    If .Range("ErrorCheckRows").EntireColumn.Hidden Then
        .Outline.ShowLevels RowLevels:=2, ColumnLevels:=2
        .Range("A1").Select
    Else
        .Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
        .Range("Heading").Offset(2, 0).Select
    End If
    
    End With
    
End Sub



Private Sub FormatReportingSheet()

    Dim wkbReportWorkbook As Workbook
    Dim wkbActivePriorToSub As Workbook

    'Necessary to activate wkbReportWorkbook to set zoom and no gridlines
    Set wkbActivePriorToSub = ActiveWorkbook
    Set wkbReportWorkbook = this.shtReport.Parent
    wkbReportWorkbook.Activate
    this.shtReport.Select

    With this.shtReport
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 11
        .DisplayPageBreaks = False
        .Columns.item("A:A").ColumnWidth = 4
        
        .Range(csHeadingAddress).Font.Bold = True
        .Range(csHeadingAddress).Font.Size = 16
        
        .Range(csCategoryAddress).Font.Color = RGB(170, 170, 170)
        .Range(csCategoryAddress).Font.Size = 8
        
        .Range(csWorkbookErrorStatus).Font.Color = RGB(170, 170, 170)
        
        'Set error check formatting
        .Range("A1").Value = "Error check workings in red gridlines"
        .Range("A1").Font.Bold = True
        FormatRedDotBorders .Range(csErrorCheckColumns)
        FormatRedDotBorders .Range(csErrorCheckRows)
        
        'Group the error checking range
        .Range(csErrorCheckColumns).EntireRow.Group
        'ActiveSheet.Outline.ShowLevels RowLevels:=1
        .Range(csErrorCheckRows).EntireColumn.Group
        ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1

        
        
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.Zoom = 80
    End With
    
    SetWorkbookErrorCellConditionalFormatting
    
    'Reactivate the original active workbook
    wkbActivePriorToSub.Activate
    
End Sub


Private Sub FormatRedDotBorders(ByVal rng As Range)

    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlDot
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlDot
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlDot
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlDot
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
        .Weight = xlThin
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


Private Function WorkbookErrorCheckFormula() As String

    WorkbookErrorCheckFormula = "=IFERROR(" & Chr(10) & _
    "      IF(" & Chr(10) & _
    "              COUNTIFS(INDIRECT(""Index!ErrorChecks""), FALSE)=0," & Chr(10) & _
    "              ""OK""," & Chr(10) & _
    "              ""Workbook error - see index tab""" & Chr(10) & _
    "           )," & Chr(10) & _
    "      ""Error checking not set""" & Chr(10) & _
    ")"

End Function


Private Sub SetWorkbookErrorCellConditionalFormatting()


    Dim rng As Range
    Dim i As Integer
    Dim ErrorFormatCondition As FormatCondition
    Dim ErrorCheckFormatConditionIsError As FormatCondition
    Dim ErrorCheckFormatConditionNotSet As FormatCondition
    
    Set rng = this.shtReport.Range("WorkbookErrorStatus")
    
    'Delete existing conditional formatting - needs to be done
    'in reverse order to prevent errors
    For i = rng.FormatConditions.Count To 1 Step -1
        rng.FormatConditions(i).Delete
    Next i
    
    Set ErrorCheckFormatConditionNotSet = rng.FormatConditions.Add( _
        Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Error checking not set""")
    With ErrorCheckFormatConditionNotSet.Font
        .Bold = True
        .Italic = False
        .Color = RGB(255, 122, 0)
        .TintAndShade = 0
    End With

    Set ErrorCheckFormatConditionIsError = rng.FormatConditions.Add( _
        Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Workbook error - see index tab""")
    With ErrorCheckFormatConditionIsError.Font
        .Bold = True
        .Italic = False
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
    End With

End Sub

