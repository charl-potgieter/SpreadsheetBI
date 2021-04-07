VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormattedSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Reporting.Sheet"
Option Explicit

Implements IFormattedSheet


Private Type TFormattedSheet
    ReportSheet As Worksheet
    Heading As String
    Category As String
    FreezePaneRow As Long
    FreezePaneCol As Long
End Type

Private this As TFormattedSheet
Private Const HeadingAddress As String = "=$B$2"
Private Const CategoryAddress As String = "=$A$1"


Public Sub Create(ByVal SelectedWorkbook As Workbook, _
    Optional ByVal SheetPosition As Long = 0)
    
    Dim NumberOfSheetsInWorkbook As Long
    
    If SheetPosition <> 0 Then
        Set this.ReportSheet = SelectedWorkbook.Sheets.Add _
            (After:=SelectedWorkbook.Worksheets.Item(SheetPosition))
    Else
        NumberOfSheetsInWorkbook = SelectedWorkbook.Sheets.Count
        Set this.ReportSheet = SelectedWorkbook.Sheets.Add _
            (After:=SelectedWorkbook.Worksheets.Item(NumberOfSheetsInWorkbook))
    End If
        
    this.FreezePaneCol = 0
    this.FreezePaneRow = 0
    
    
End Sub

Private Property Get IFormattedSheet_Name() As String
    IFormattedSheet_Name = this.ReportSheet.Name
End Property


Private Property Let IFormattedSheet_Name(ByVal Name As String)

    Dim Counter As Long
    Dim TryName As String
    Counter = 0
    
    'Get first availabe sheet name add numeric prefix if name is taken
    TryName = Name
    On Error Resume Next
    Do While this.ReportSheet.Name <> TryName
        If Counter <> 0 Then
            TryName = Name & "_" & Counter
        End If
        this.ReportSheet.Name = TryName
         Counter = Counter + 1
    Loop
    On Error GoTo 0
        
End Property


Private Property Get IFormattedSheet_Heading() As String
    IFormattedSheet_Heading = this.Heading
End Property

Private Property Let IFormattedSheet_Heading(ByVal Heading As String)
    
    this.Heading = Heading
    With this.ReportSheet
        On Error Resume Next
        .Names.Item("Heading").Delete
        On Error GoTo 0
        .Names.Add Name:="Heading", RefersTo:="=$B$2"
        .Range("Heading").Value = Heading
    End With
    
End Property

Private Property Get IFormattedSheet_Category() As String
    IFormattedSheet_Category = this.Category
End Property

Private Property Let IFormattedSheet_Category(ByVal Category As String)
    
    this.Category = Category
    With this.ReportSheet
        On Error Resume Next
        .Names.Item("Category").Delete
        On Error GoTo 0
        .Names.Add Name:="Category", RefersTo:=CategoryAddress
        .Range("Category").Value = Category
    End With
    
End Property


Private Property Get IFormattedSheet_FreezePaneRow() As Long
    IFormattedSheet_FreezePaneRow = this.FreezePaneRow
End Property


Private Property Let IFormattedSheet_FreezePaneRow(ByVal FreezeRow As Long)
    this.FreezePaneRow = FreezeRow
    FreezePanes
End Property


Private Property Get IFormattedSheet_FreezePaneCol() As Long
    IFormattedSheet_FreezePaneCol = this.FreezePaneCol
End Property


Private Property Let IFormattedSheet_FreezePaneCol(ByVal FreezeColumn As Long)
    this.FreezePaneCol = FreezeColumn
    FreezePanes
End Property


Private Sub IFormattedSheet_Format()

    Dim ReportWorkbook As Workbook
    Dim ActiveWorkbookPriorToSub As Workbook

    'Necessary to activate Reportworkbook to set zoom and no gridlines
    Set ActiveWorkbookPriorToSub = ActiveWorkbook
    Set ReportWorkbook = this.ReportSheet.Parent
    ReportWorkbook.Activate
    this.ReportSheet.Select

    With this.ReportSheet
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 11
        .DisplayPageBreaks = False
        .Columns.Item("A:A").ColumnWidth = 4
        .Range(HeadingAddress).Font.Bold = True
        .Range(HeadingAddress).Font.Size = 16
        .Range(CategoryAddress).Font.Color = RGB(170, 170, 170)
        .Range(CategoryAddress).Font.Size = 8
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.Zoom = 80
    End With
    
    'Reactivate the original active workbook
    ActiveWorkbookPriorToSub.Activate
    
End Sub



Private Sub FreezePanes()
   
    Dim WorkbookSelectedBeforeSub As Workbook
    Dim SelectedSheetOnReportingWorkbookBeforeSub As Worksheet
   ' Dim SelectionOnReportingSheetBeforeSub As Range
    
    Set WorkbookSelectedBeforeSub = ActiveWorkbook
    this.ReportSheet.Parent.Activate
    Set SelectedSheetOnReportingWorkbookBeforeSub = ActiveSheet
    this.ReportSheet.Activate
    
    'SelectionOnReportingSheetBeforeSub = Selection
    With ActiveWindow
        .FreezePanes = False
        .Split = False
        Select Case True
        Case this.FreezePaneCol = 0 And this.FreezePaneRow <> 0
            .SplitRow = this.FreezePaneRow
            .FreezePanes = True
        Case this.FreezePaneCol <> 0 And this.FreezePaneRow = 0
            .SplitColumn = this.FreezePaneCol
            .FreezePanes = True
        Case this.FreezePaneCol <> 0 And this.FreezePaneRow <> 0
            .SplitColumn = this.FreezePaneCol
            .SplitRow = this.FreezePaneRow
            .FreezePanes = True
        End Select
    End With
    
    SelectedSheetOnReportingWorkbookBeforeSub.Activate
    WorkbookSelectedBeforeSub.Activate
    
End Sub


