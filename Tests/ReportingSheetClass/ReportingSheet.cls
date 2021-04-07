VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReportingSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Reporting"
Option Explicit

Private Type TReportingSheet
    shtReport As Worksheet
    Heading As String
    Category As String
    FreezePaneRow As Long
    FreezePaneCol As Long
    tbl As ReportingTable
End Type

Private this As TReportingSheet
Private Const csHeadingAddress As String = "=$B$2"
Private Const csCategoryAddress As String = "=$A$1"


Public Sub Create(ByVal wkb As Workbook, _
    Optional ByVal SheetPosition As Long = 0)
    
    Dim NumberOfSheetsInWorkbook As Long
    
    If SheetPosition <> 0 Then
        Set this.shtReport = wkb.Sheets.Add _
            (After:=wkb.Worksheets.Item(SheetPosition))
    Else
        NumberOfSheetsInWorkbook = wkb.Sheets.Count
        Set this.shtReport = wkb.Sheets.Add _
            (After:=wkb.Worksheets.Item(NumberOfSheetsInWorkbook))
    End If
        
    this.FreezePaneCol = 0
    this.FreezePaneRow = 0
    
    
End Sub


Public Function ImportExistingSheet(ByVal sht As Worksheet) As Boolean
    
    Select Case True
    
        Case Not this.shtReport Is Nothing
            ImportExistingSheet = False
    
        Case sht.PivotTables.Count > 1
            ImportExistingSheet = False
        
        Case sht.ListObjects.Count > 1
            ImportExistingSheet = False
            
        Case Else
            ImportExistingSheet = True
            Set this.shtReport = sht
            
    End Select

End Function

Public Property Get Sheet() As Worksheet
    Set Sheet = this.shtReport
End Property


Public Property Get Name() As String
    Name = this.shtReport.Name
End Property


Public Property Let Name(ByVal sName As String)

    Dim Counter As Long
    Dim TryName As String
    Counter = 0
    
    'Get first availabe sheet name add numeric prefix if name is taken
    TryName = sName
    On Error Resume Next
    Do While this.shtReport.Name <> TryName
        If Counter <> 0 Then
            TryName = sName & "_" & Counter
        End If
        this.shtReport.Name = TryName
         Counter = Counter + 1
    Loop
    On Error GoTo 0
        
End Property


Public Property Get Heading() As String
    Heading = this.Heading
End Property

Public Property Let Heading(ByVal sHeading As String)
    
    this.Heading = sHeading
    With this.shtReport
        On Error Resume Next
        .Names.Item("Heading").Delete
        On Error GoTo 0
        .Names.Add Name:="Heading", RefersTo:="=$B$2"
        .Range("Heading").Value = Heading
    End With
    
End Property

Public Property Get Category() As String
    Category = this.Category
End Property

Public Property Let Category(ByVal sCategory As String)
    
    this.Category = sCategory
    With this.shtReport
        On Error Resume Next
        .Names.Item("Category").Delete
        On Error GoTo 0
        .Names.Add Name:="Category", RefersTo:=csCategoryAddress
        .Range("Category").Value = Category
    End With
    
End Property


Public Property Get FreezePaneRow() As Long
    FreezePaneRow = this.FreezePaneRow
End Property


Public Property Let FreezePaneRow(ByVal lFreezeRow As Long)
    this.FreezePaneRow = lFreezeRow
    FreezePanes
End Property


Public Property Get FreezePaneCol() As Long
    FreezePaneCol = this.FreezePaneCol
End Property


Public Property Let FreezePaneCol(ByVal lFreezeColumn As Long)
    this.FreezePaneCol = lFreezeColumn
    FreezePanes
End Property

Public Property Get ReportingTable() As ReportingTable
    ReportingTable = this.tbl
End Property


Public Sub Format()

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
        .Columns.Item("A:A").ColumnWidth = 4
        .Range(csHeadingAddress).Font.Bold = True
        .Range(csHeadingAddress).Font.Size = 16
        .Range(csCategoryAddress).Font.Color = RGB(170, 170, 170)
        .Range(csCategoryAddress).Font.Size = 8
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.Zoom = 80
    End With
    
    'Reactivate the original active workbook
    wkbActivePriorToSub.Activate
    
End Sub


Public Function AddReportingTable(ByVal rng As Range) As Boolean

    Select Case True

        Case Intersect(rng, this.shtReport.Cells).Cells.Count <> rng.Cells.Count
            'Return false if rng is not contained within the ReportingSheet
            AddReportingTable = False
        
        Case Not this.tbl Is Nothing
            AddReportingTable = False
            
        Case Else
            AddReportingTable = True
            Set this.tbl = New ReportingTable
            this.tbl.Create rng
    
    End Select

End Function

Private Sub FreezePanes()
   
    Dim WorkbookSelectedBeforeSub As Workbook
    Dim SelectedSheetOnReportingWorkbookBeforeSub As Worksheet
   ' Dim SelectionOnReportingSheetBeforeSub As Range
    
    Set WorkbookSelectedBeforeSub = ActiveWorkbook
    this.shtReport.Parent.Activate
    Set SelectedSheetOnReportingWorkbookBeforeSub = ActiveSheet
    this.shtReport.Activate
    
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


