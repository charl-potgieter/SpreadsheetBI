VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IndexPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TypeIndedxSheet
    IndexSheet As Worksheet
End Type

Private this As TypeIndedxSheet


Public Sub Create(ByVal wkb As Workbook)

    On Error Resume Next
    wkb.Sheets("Index").Delete
    On Error GoTo 0
    
    Set this.IndexSheet = wkb.Worksheets.Add(Before:=wkb.Sheets(1))
    this.IndexSheet.Name = "Index"

    SetIndexSheetRangeNames
    FormatIndexSheet
    CreateIndexSheetFormulas
    CreateLinkagesBetweenIndexAndTargetPages
    

End Sub



Private Sub SetIndexSheetRangeNames()

    With this.IndexSheet.Names
        .Add Name:="HiddenSheetNamesCol", RefersToLocal:=Range("$A:$A")
        .Add Name:="HiddenCategoriesCol", RefersToLocal:=Range("$B:$B")
        .Add Name:="CategoryCol", RefersToLocal:=Range("$D:D")
        .Add Name:="ReportNamesCol", RefersToLocal:=Range("$E:$E")
        .Add Name:="ErrorCheckCol", RefersToLocal:=Range("$F:$F")
        .Add Name:="SheetHeading", RefersToLocal:=Range("$D$2")
        .Add Name:="DefaultCursorLocation", RefersToLocal:=Range("$D$4")
    End With

End Sub


Private Sub FormatIndexSheet()

    With this.IndexSheet

        .Activate
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 11
        .DisplayPageBreaks = False

        .Range("C:C").EntireColumn.ColumnWidth = 4
        .Range("ErrorCheckCol").EntireColumn.ColumnWidth = 13
        .Range("ReportNamesCol").EntireColumn.ColumnWidth = 100
        .Range("HiddenSheetNamesCol").EntireColumn.ColumnWidth = 30
        .Range("HiddenCategoriesCol").EntireColumn.ColumnWidth = 30
        .Range("HiddenSheetNamesCol").Interior.Color = RGB(217, 217, 217)
        .Range("HiddenCategoriesCol").Interior.Color = RGB(217, 217, 217)
        

        With Union(.Range("HiddenSheetNamesCol"), .Range("HiddenCategoriesCol")).Columns
            .Group
            .Resize(, 1).EntireColumn.ShowDetail = False
        End With

        .Range("A2").Value = "Hidden workings in grey"

        .Range("CategoryCol").Font.Bold = True

        .Range("SheetHeading").Font.Bold = True
        .Range("SheetHeading").Font.Size = 16
        .Range("SheetHeading").Value = "Index"

        .Range("ErrorCheckCol").Cells(3) = "Errors OK?"
        .Range("ErrorCheckCol").Cells(3).Font.Bold = True

        .Range("HiddenSheetNamesCol").Cells(5) = "Sheet Name"
        .Range("HiddenCategoriesCol").Cells(5) = "Category"
        .Range("HiddenSheetNamesCol").Cells(5).Font.Bold = True
        .Range("HiddenCategoriesCol").Cells(5).Font.Bold = True

        .Rows("4:4").Select
        ActiveWindow.FreezePanes = True
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.Zoom = 80

    End With

End Sub


Private Sub CreateIndexSheetFormulas()

    Dim ErrorCheckFormatCondition As FormatCondition
    Dim Temp As String

    With this.IndexSheet.Range("CategoryCol").Cells(5)
        .Value = "No category duplicates (duplicates indicate out of order sheets)"
        .Font.Color = RGB(170, 170, 170)
        .Font.Bold = False
    End With

    With this.IndexSheet.Range("ErrorCheckCol").Cells(5)
        .Formula = _
            "=COUNTA(" & vbLf & _
            "    FILTER(CategoryCol, NOT(ISBLANK(CategoryCol)))" & _
            ") " & vbLf & _
            " = " & vbLf & _
            " COUNTA(" & vbLf & _
            "    UNIQUE(FILTER(CategoryCol, NOT(ISBLANK(CategoryCol))))" & vbLf & _
            ")"
        .Font.Color = RGB(170, 170, 170)
        Set ErrorCheckFormatCondition = .FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE")
        ErrorCheckFormatCondition.Font.Bold = True
        ErrorCheckFormatCondition.Font.Color = RGB(255, 0, 0)
    End With

    With this.IndexSheet.Range("CategoryCol").Cells(6)
        .Value = "No duplicate category / report name combinations"
        .Font.Color = RGB(170, 170, 170)
        .Font.Bold = False
    End With


    With this.IndexSheet.Range("ErrorCheckCol").Cells(6)
        .Formula = _
            "=COUNTA(" & vbLf & _
            "    FILTER(HiddenCategoriesCol & ReportNamesCol, NOT(ISBLANK(ReportNamesCol)))" & vbLf & _
            ")" & vbLf & _
            " =" & vbLf & _
            "COUNTA(" & vbLf & _
            "    UNIQUE(Filter(HiddenCategoriesCol & ReportNamesCol, NOT(ISBLANK(ReportNamesCol))))" & vbLf & _
            ")"
        .Font.Color = RGB(170, 170, 170)
        Set ErrorCheckFormatCondition = .FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE")
        ErrorCheckFormatCondition.Font.Bold = True
        ErrorCheckFormatCondition.Font.Color = RGB(255, 0, 0)
    End With

    Set ErrorCheckFormatCondition = Nothing

End Sub



Private Sub CreateLinkagesBetweenIndexAndTargetPages()

    Dim sht As Worksheet
    Dim wkb As Workbook
    Dim CurrentRow As Long
    Dim LastCapturedReportCategory As String
    Dim ReportSheet As ReportingSheet
    Dim Storage As ListStorage
    Dim ReportSheetAssigned As Boolean
    Dim StorageAssigned As Boolean
    Dim ReportNameCellAddress As String
    
    Set wkb = this.IndexSheet.Parent
    CurrentRow = 5
    LastCapturedReportCategory = ""

    For Each sht In wkb.Worksheets
        Set ReportSheet = New ReportingSheet
        ReportSheetAssigned = ReportSheet.AssignExistingSheet(sht)
        If ReportSheetAssigned And (sht.Visible = xlSheetVisible) Then
            
            WriteCategoryNameOnIndexPage ReportSheet.Category, CurrentRow, LastCapturedReportCategory
            WriteReportNameOnIndexPage ReportSheet.Heading, ReportSheet.Name, CurrentRow
            WriteHiddenSheetNameOnIndexPage ReportSheet.Name, CurrentRow
            WriteHiddenCategoryNameOnIndexPage ReportSheet.Category, CurrentRow
            WriteReferenceToSheetErrorCheckOnIndexPage ReportSheet, CurrentRow
            ReportSheet.RefreshIndexSheetRelatedItems
            
        Else
            Set Storage = New ListStorage
            StorageAssigned = Storage.AssignStorage(sht.Parent, sht.Name)
            If StorageAssigned Then

                WriteCategoryNameOnIndexPage "ListStorage", CurrentRow, LastCapturedReportCategory
                WriteReportNameOnIndexPage Storage.Name, Storage.Name, CurrentRow
                WriteHiddenSheetNameOnIndexPage Storage.Name, CurrentRow
                WriteHiddenCategoryNameOnIndexPage "ListStorage", CurrentRow
            

            End If
        End If
        Set ReportSheet = Nothing
    Next sht


End Sub






Sub WriteCategoryNameOnIndexPage(ByVal SheetCategory As String, _
    ByRef CurrentRow As Long, ByRef LastCapturedReportCategory As String)

    If SheetCategory <> LastCapturedReportCategory Then
        CurrentRow = CurrentRow + 3
        LastCapturedReportCategory = SheetCategory
        this.IndexSheet.Range("CategoryCol").Cells(CurrentRow) = SheetCategory
    End If

End Sub


Sub WriteReportNameOnIndexPage(ByVal SheetHeading As String, _
    ByVal SheetName As String, ByRef CurrentRow As Long)

    CurrentRow = CurrentRow + 2
    this.IndexSheet.Range("ReportNamesCol").Cells(CurrentRow) = SheetHeading
    ActiveSheet.Hyperlinks.Add _
        Anchor:=this.IndexSheet.Range("ReportNamesCol").Cells(CurrentRow), _
        Address:="", _
        SubAddress:="'" & SheetName & "'" & "!DefaultCursorLocation"

End Sub


Sub WriteHiddenSheetNameOnIndexPage(ByVal SheetName As String, _
    ByRef CurrentRow As Long)

    'Write sheet name in hidden column (not used anymore but could change)
    this.IndexSheet.Range("HiddenSheetNamesCol").Cells(CurrentRow) = SheetName

End Sub


Sub WriteHiddenCategoryNameOnIndexPage(ByVal SheetCategory As String, _
    ByRef CurrentRow As Long)

    'Write sheet name in hidden column (not used anymore but could change)
    this.IndexSheet.Range("HiddenCategoriesCol").Cells(CurrentRow) = SheetCategory

End Sub




Sub WriteReferenceToSheetErrorCheckOnIndexPage(ByVal ReportSheet As ReportingSheet, _
    ByVal CurrentRow As Long)

    Dim IndexPageErrorCheckFormula As String
    Dim ErrorCheckFormatCondition As FormatCondition


    IndexPageErrorCheckFormula = _
        "=IFERROR(" & vbLf & _
        "   " & ReportSheet.ErrorStatusRangeAddress & "=""OK""," & vbLf & _
        "   FALSE" & vbLf & _
        ")"

    With this.IndexSheet.Range("ErrorCheckCol")
        .Cells(CurrentRow).Formula = IndexPageErrorCheckFormula
        .Cells(CurrentRow).Font.Color = RGB(170, 170, 170)
        Set ErrorCheckFormatCondition = .Cells(CurrentRow).FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE")
        ErrorCheckFormatCondition.Font.Bold = True
        ErrorCheckFormatCondition.Font.Color = RGB(255, 0, 0)
    End With

    Set ErrorCheckFormatCondition = Nothing

End Sub



Sub ActivateAtDefauiltCursorLocation()
    this.IndexSheet.Activate
    this.IndexSheet.Range("DefaultCursorLocation").Select
End Sub



Private Sub Class_Terminate()
    Set this.IndexSheet = Nothing
End Sub






