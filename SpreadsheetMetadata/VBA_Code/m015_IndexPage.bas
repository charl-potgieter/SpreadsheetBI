Attribute VB_Name = "m015_IndexPage"
Option Explicit


Function InsertIndexPage(ByVal wkb As Workbook) As Worksheet

    Dim IndexSheet As Worksheet
    Dim sht As Worksheet
    Dim ReportSheet As ReportingSheet
    Dim LastCapturedReportCategory As String
    Dim CurrentRow As Long
    Dim ReportSheetAssigned As Boolean
    
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
            CreateReturnToIndexLink ReportSheet
            WriteCategoryName IndexSheet, ReportSheet, CurrentRow, LastCapturedReportCategory
            WriteReportName IndexSheet, ReportSheet, CurrentRow
            WriteHiddenSheetName IndexSheet, ReportSheet, CurrentRow
            WriteHiddenCategoryName IndexSheet, ReportSheet, CurrentRow
            WriteReferenceToSheetErrorCheck IndexSheet, ReportSheet, CurrentRow
            ReportSheet.WorkbookErrorStatusFormula = WorkbookErrorStatusFormula
            ReportSheet.SheetErrorStatusFormula = SheetErrorStatusFormula
        End If
    Next sht
    
    IndexSheet.Activate
    IndexSheet.Range("DefaultCursorLocation").Select
    Set InsertIndexPage = IndexSheet

End Function


Private Function CreateIndexSheet(ByVal wkb As Workbook) As Worksheet

    Dim sht As Worksheet
    
    On Error Resume Next
    wkb.Sheets("Index").Delete
    On Error GoTo 0
    Set sht = wkb.Sheets.Add(Before:=wkb.Sheets(1))
    sht.Name = "Index"
    Set CreateIndexSheet = sht

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
    Dim temp As String

    With IndexSheet.Range("CategoryCol").Cells(5)
        .Value = "No category duplicates (duplicates indicate out of order sheets)"
        .Font.Color = rgb(170, 170, 170)
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
        .Font.Color = rgb(170, 170, 170)
        Set ErrorCheckFormatCondition = .FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE")
        ErrorCheckFormatCondition.Font.Bold = True
        ErrorCheckFormatCondition.Font.Color = rgb(255, 0, 0)
    End With

    With IndexSheet.Range("CategoryCol").Cells(6)
        .Value = "No duplicate category / report name combinations"
        .Font.Color = rgb(170, 170, 170)
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
        .Font.Color = rgb(170, 170, 170)
        Set ErrorCheckFormatCondition = .FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE")
        ErrorCheckFormatCondition.Font.Bold = True
        ErrorCheckFormatCondition.Font.Color = rgb(255, 0, 0)
    End With
    


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

End Sub



Sub CreateReturnToIndexLink(ByVal ReportSheet As ReportingSheet)

    ReportSheet.Sheet.Hyperlinks.Add _
        Anchor:=ReportSheet.Sheet.Range("ReturnToIndex"), _
        Address:="", _
        SubAddress:="Index!DefaultCursorLocation", _
        TextToDisplay:="<Return to Index>"

End Sub


Sub WriteCategoryName(ByVal IndexSheet As Worksheet, ByVal ReportSheet As ReportingSheet, _
    ByRef CurrentRow As Long, ByRef LastCapturedReportCategory As String)

    If ReportSheet.Category <> LastCapturedReportCategory Then
        CurrentRow = CurrentRow + 3
        LastCapturedReportCategory = ReportSheet.Category
        IndexSheet.Range("CategoryCol").Cells(CurrentRow) = ReportSheet.Category
    End If

End Sub


Sub WriteReportName(ByVal IndexSheet As Worksheet, ByVal ReportSheet As ReportingSheet, _
    ByRef CurrentRow As Long)

    CurrentRow = CurrentRow + 2
    IndexSheet.Range("ReportNamesCol").Cells(CurrentRow) = ReportSheet.Heading
    ActiveSheet.Hyperlinks.Add _
        Anchor:=IndexSheet.Range("ReportNamesCol").Cells(CurrentRow), _
        Address:="", _
        SubAddress:="'" & ReportSheet.Sheet.Name & "'" & "!DefaultCursorLocation"

End Sub


Sub WriteHiddenSheetName(ByVal IndexSheet As Worksheet, ByVal ReportSheet As ReportingSheet, _
    ByRef CurrentRow As Long)

    'Write sheet name in hidden column (not used anymore but could change)
    IndexSheet.Range("HiddenSheetNamesCol").Cells(CurrentRow) = ReportSheet.Sheet.Name

End Sub


Sub WriteHiddenCategoryName(ByVal IndexSheet As Worksheet, ByVal ReportSheet As ReportingSheet, _
    ByRef CurrentRow As Long)

    'Write sheet name in hidden column (not used anymore but could change)
    IndexSheet.Range("HiddenCategoriesCol").Cells(CurrentRow) = ReportSheet.Category

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
        .Cells(CurrentRow).Font.Color = rgb(170, 170, 170)
        Set ErrorCheckFormatCondition = .Cells(CurrentRow).FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE")
        ErrorCheckFormatCondition.Font.Bold = True
        ErrorCheckFormatCondition.Font.Color = rgb(255, 0, 0)
    End With

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






