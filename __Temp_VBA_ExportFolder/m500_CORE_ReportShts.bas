Attribute VB_Name = "m500_CORE_ReportShts"
Option Explicit
Option Private Module


Function InsertIndexPage(ByVal wkb As Workbook) As Worksheet

    Dim IndexSheet As Worksheet
    Dim sht As Worksheet
    Dim ReportSheet As ReportingSheet
    Dim Storage As ListStorage
    Dim LastCapturedReportCategory As String
    Dim CurrentRow As Long
    Dim ReportSheetAssigned As Boolean
    Dim StorageAssigned As Boolean
    
    Set IndexSheet = CreateIndexSheet(wkb)
    SetIndexSheetRangeNames IndexSheet
    FormatIndexSheet IndexSheet
    CreateIndexSheetFormulas IndexSheet
    AddFirstAndLastSheets wkb
    CurrentRow = 5
    LastCapturedReportCategory = ""

    For Each sht In wkb.Worksheets
        Set ReportSheet = New ReportingSheet
        ReportSheetAssigned = ReportSheet.AssignExistingSheet(sht)
        If ReportSheetAssigned And (sht.Visible = xlSheetVisible) Then
            CreateReturnToIndexLink ReportSheet.Sheet
            WriteCategoryName IndexSheet, ReportSheet.Category, CurrentRow, LastCapturedReportCategory
            WriteReportName IndexSheet, ReportSheet.Heading, ReportSheet.Name, CurrentRow
            WriteHiddenSheetName IndexSheet, ReportSheet.Name, CurrentRow
            WriteHiddenCategoryName IndexSheet, ReportSheet.Category, CurrentRow
            WriteReferenceToSheetErrorCheck IndexSheet, ReportSheet, CurrentRow
            ReportSheet.WorkbookErrorStatusFormula = WorkbookErrorStatusFormula
            ReportSheet.SheetErrorStatusFormula = SheetErrorStatusFormula
        Else
            Set Storage = New ListStorage
            StorageAssigned = Storage.AssignStorage(sht.Parent, sht.Name)
            If StorageAssigned Then
                CreateReturnToIndexLink Storage.Sheet
                WriteCategoryName IndexSheet, "ListStorage", CurrentRow, LastCapturedReportCategory
                WriteReportName IndexSheet, Storage.Name, Storage.Name, CurrentRow
                WriteHiddenSheetName IndexSheet, Storage.Name, CurrentRow
                WriteHiddenCategoryName IndexSheet, "ListStorage", CurrentRow
            End If
        End If
        Set ReportSheet = Nothing
    Next sht
    
    IndexSheet.Activate
    IndexSheet.Range("DefaultCursorLocation").Select
    Set InsertIndexPage = IndexSheet

    Set IndexSheet = Nothing
    Set sht = Nothing
    Set ReportSheet = Nothing

End Function


Private Function CreateIndexSheet(ByVal wkb As Workbook) As Worksheet

    Dim sht As Worksheet
    
    On Error Resume Next
    wkb.Sheets("Index").Delete
    On Error GoTo 0
    Set sht = wkb.Sheets.Add(Before:=wkb.Sheets(1))
    sht.Name = "Index"
    Set CreateIndexSheet = sht
    Set sht = Nothing

End Function


Private Sub SetIndexSheetRangeNames(ByVal sht As Worksheet)

    With sht.Names
        .Add Name:="HiddenSheetNamesCol", RefersToLocal:=Range("$A:$A")
        .Add Name:="HiddenCategoriesCol", RefersToLocal:=Range("$B:$B")
        .Add Name:="CategoryCol", RefersToLocal:=Range("$D:D")
        .Add Name:="ReportNamesCol", RefersToLocal:=Range("$E:$E")
        .Add Name:="ErrorCheckCol", RefersToLocal:=Range("$F:$F")
        .Add Name:="SheetHeading", RefersToLocal:=Range("$D$2")
        .Add Name:="DefaultCursorLocation", RefersToLocal:=Range("$D$4")
    End With

End Sub


Private Sub FormatIndexSheet(ByRef sht As Worksheet)

    With sht
    
        .Activate
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 11
        .DisplayPageBreaks = False
        
        .Range("C:C").EntireColumn.ColumnWidth = 4
        .Range("ErrorCheckCol").EntireColumn.ColumnWidth = 13
        .Range("ReportNamesCol").EntireColumn.ColumnWidth = 100
        .Range("HiddenSheetNamesCol").EntireColumn.ColumnWidth = 30
        .Range("HiddenCategoriesCol").EntireColumn.ColumnWidth = 30
        .Range("HiddenSheetNamesCol").EntireColumn.Hidden = True
        .Range("HiddenCategoriesCol").EntireColumn.Hidden = True
        
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

Sub CreateIndexSheetFormulas(ByVal IndexSheet As Worksheet)

    Dim ErrorCheckFormatCondition As FormatCondition
    Dim Temp As String

    With IndexSheet.Range("CategoryCol").Cells(5)
        .Value = "No category duplicates (duplicates indicate out of order sheets)"
        .Font.Color = RGB(170, 170, 170)
        .Font.Bold = False
    End With

    With IndexSheet.Range("ErrorCheckCol").Cells(5)
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

    With IndexSheet.Range("CategoryCol").Cells(6)
        .Value = "No duplicate category / report name combinations"
        .Font.Color = RGB(170, 170, 170)
        .Font.Bold = False
    End With
    
    
    With IndexSheet.Range("ErrorCheckCol").Cells(6)
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



Private Sub AddFirstAndLastSheets(ByVal wkb As Workbook)

    Dim FirstSheet As Worksheet
    Dim LastSheet As Worksheet

    'Create an empty hidden first and last sheet as anchor points for 3d sum range
    'for storing sheet hashes to check completeness of index page
    On Error Resume Next
    wkb.Sheets("FirstSheet").Delete
    wkb.Sheets("LastSheet").Delete
    On Error GoTo 0
    
    Set FirstSheet = wkb.Sheets.Add(Before:=wkb.Sheets(1))
    Set LastSheet = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FirstSheet.Name = "FirstSheet"
    LastSheet.Name = "LastSheet"
    FirstSheet.Visible = xlSheetHidden
    LastSheet.Visible = xlSheetHidden

    Set FirstSheet = Nothing
    Set LastSheet = Nothing

End Sub



Sub CreateReturnToIndexLink(ByVal sht As Worksheet)
    
    Dim ReturnToIndexShape As Shape
    
    On Error Resume Next
    sht.Shapes("ReturnToIndex").Delete
    On Error GoTo 0

    Set ReturnToIndexShape = sht.Shapes.AddShape( _
        msoShapeRoundedRectangle, sht.Range("ReturnToIndex").Left, _
        sht.Range("ReturnToIndex").Offset(1, 0).Top, _
        100, 21)
    
    ReturnToIndexShape.Name = "ReturnToIndex"
    
    With ReturnToIndexShape.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.RGB = RGB(240, 240, 240)
        .Transparency = 0
        .Solid
    End With
    
    With ReturnToIndexShape.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Weight = 0.5
    End With
    
    With ReturnToIndexShape.TextFrame2.TextRange
        .Text = "Return to index"
        .Font.Fill.Visible = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Font.Fill.Transparency = 0
        .Font.Fill.Solid
        .ParagraphFormat.FirstLineIndent = 0
        .ParagraphFormat.Alignment = msoAlignLeft
        .Font.Fill.Visible = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Font.Fill.Transparency = 0
        .Font.Fill.Solid
        .Font.Size = 9
    End With
    
    ReturnToIndexShape.Placement = xlFreeFloating
    
    sht.Hyperlinks.Add Anchor:=ReturnToIndexShape, _
        Address:="", _
        SubAddress:="Index!DefaultCursorLocation"

    Set ReturnToIndexShape = Nothing
    Set sht = Nothing

End Sub


Sub WriteCategoryName(ByVal IndexSheet As Worksheet, ByVal SheetCategory As String, _
    ByRef CurrentRow As Long, ByRef LastCapturedReportCategory As String)

    If SheetCategory <> LastCapturedReportCategory Then
        CurrentRow = CurrentRow + 3
        LastCapturedReportCategory = SheetCategory
        IndexSheet.Range("CategoryCol").Cells(CurrentRow) = SheetCategory
    End If

End Sub


Sub WriteReportName(ByVal IndexSheet As Worksheet, ByVal SheetHeading As String, _
    ByVal SheetName As String, ByRef CurrentRow As Long)

    CurrentRow = CurrentRow + 2
    IndexSheet.Range("ReportNamesCol").Cells(CurrentRow) = SheetHeading
    ActiveSheet.Hyperlinks.Add _
        Anchor:=IndexSheet.Range("ReportNamesCol").Cells(CurrentRow), _
        Address:="", _
        SubAddress:="'" & SheetName & "'" & "!DefaultCursorLocation"

End Sub


Sub WriteHiddenSheetName(ByVal IndexSheet As Worksheet, ByVal SheetName As String, _
    ByRef CurrentRow As Long)

    'Write sheet name in hidden column (not used anymore but could change)
    IndexSheet.Range("HiddenSheetNamesCol").Cells(CurrentRow) = SheetName

End Sub


Sub WriteHiddenCategoryName(ByVal IndexSheet As Worksheet, ByVal SheetCategory As String, _
    ByRef CurrentRow As Long)

    'Write sheet name in hidden column (not used anymore but could change)
    IndexSheet.Range("HiddenCategoriesCol").Cells(CurrentRow) = SheetCategory

End Sub


Sub WriteReferenceToSheetErrorCheck(ByVal IndexSheet As Worksheet, _
    ByVal ReportSheet As ReportingSheet, ByVal CurrentRow As Long)

    Dim SheetErrorStatusRangeName As String
    Dim IndexPageErrorCheckFormula As String
    Dim ErrorCheckFormatCondition As FormatCondition
    
    
    SheetErrorStatusRangeName = "'" & ReportSheet.Name & "'!SheetErrorStatus"
    
    
    IndexPageErrorCheckFormula = _
        "=IFERROR(" & vbLf & _
        "   " & SheetErrorStatusRangeName & "=""OK""," & vbLf & _
        "   FALSE" & vbLf & _
        ")"
    
    With IndexSheet.Range("ErrorCheckCol")
        .Cells(CurrentRow).Formula = IndexPageErrorCheckFormula
        .Cells(CurrentRow).Font.Color = RGB(170, 170, 170)
        Set ErrorCheckFormatCondition = .Cells(CurrentRow).FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE")
        ErrorCheckFormatCondition.Font.Bold = True
        ErrorCheckFormatCondition.Font.Color = RGB(255, 0, 0)
    End With

    Set ErrorCheckFormatCondition = Nothing

End Sub



Private Function WorkbookErrorStatusFormula() As String

    WorkbookErrorStatusFormula = _
        "=IFERROR(" & vbLf & _
        "   IF(" & vbLf & _
        "       COUNTIFS(Index!ErrorCheckCol, FALSE) <> 0," & vbLf & _
        "      ""Workbook error - see index page""," & vbLf & _
        "      ""OK""" & vbLf & _
        "   )," & vbLf & _
        "   ""Workbook error - see index page""" & vbLf & _
        ")"

End Function


Private Function SheetErrorStatusFormula() As String

    SheetErrorStatusFormula = _
        "=IFERROR(" & vbLf & _
        "   SWITCH(" & vbLf & _
        "       TRUE," & vbLf & _
        "       NOT(AND(" & vbLf & _
        "           COUNTIFS(ErrorCheckColumns, FALSE) = 0," & vbLf & _
        "           COUNTIFS(ErrorCheckRows, FALSE) = 0," & vbLf & _
        "           SUMPRODUCT(--ISERROR(ErrorCheckColumns))=0," & vbLf & _
        "           SumProduct(--IsError(ErrorCheckRows)) = 0" & vbLf & _
        "       )), ""Sheet error check issue - see ranges ErrrorCheckColumns and ErrorCheckRows""," & vbLf & _
        "       COUNTIFS(Index!HiddenCategoriesCol, Category, Index!ReportNamesCol, Heading) = 0, ""This sheet heading / category combination does not appear on index tab""," & vbLf & _
        "       COUNTIFS(Index!HiddenCategoriesCol, Category, Index!ReportNamesCol, Heading) > 1, ""This sheet heading / category combination appears multiple times on index tab""," & vbLf & _
        "       ""OK""" & vbLf & _
        "   )," & vbLf & _
        "   ""Sheet error""" & vbLf & _
        ")"


End Function

