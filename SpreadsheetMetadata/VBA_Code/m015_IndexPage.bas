Attribute VB_Name = "m015_IndexPage"
Option Explicit


Sub InsertIndexPage(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim ReportSheet As ReportingSheet
    Dim bReportSheetAssigned As Boolean
    Dim shtIndex As Worksheet
    Dim i As Double
    Dim sLastCapturedReportCategory As String
    Dim rngCategoryCol As Range
    Dim rngReportCol As Range
    Dim rngSheetNameCol As Range
    Dim rngShowRange As Range
    Dim rngErrorCol As Range
    Dim sSheetErrorCheckColumnsRangeName As String
    Dim sSheetErrorCheckRowsRangeName As String
    Dim sIndexPageErrorCheckFormula As String
    Dim ErrorCheckFormatCondition As FormatCondition
    Dim FirstSheet As Worksheet
    Dim LastSheet As Worksheet
    
    'Delete any previous index sheet and create a new one
    On Error Resume Next
    wkb.Sheets("Index").Delete
    wkb.Sheets("FirstSheet").Delete
    wkb.Sheets("LastSheet").Delete
    On Error GoTo 0
    Set shtIndex = wkb.Sheets.Add(Before:=ActiveWorkbook.Sheets(1))
    FormatSheet shtIndex
    
    wkb.Activate
    shtIndex.Activate
    
    With shtIndex
    
        .Name = "Index"
        .Range("A:A").Insert Shift:=xlToRight
        .Range("A:A").EntireColumn.Hidden = True
        .Range("C2") = "Index"
        .Range("D5").Font.Bold = True
        .Range("E3") = "Errors OK?"
        .Names.Add Name:="ErrorChecks", RefersTo:="=$E:$E"
        .Range("E3").Font.Bold = True
        .Columns("E:E").ColumnWidth = 13
        .Columns("D:D").ColumnWidth = 100
        .Columns("E:E").ColumnWidth = 13
        .rows("4:4").Select
        ActiveWindow.FreezePanes = True
        
        Set rngSheetNameCol = .Columns("A")
        Set rngCategoryCol = .Columns("C")
        Set rngReportCol = .Columns("D")
        Set rngErrorCol = .Columns("E")
       
        i = 2
        sLastCapturedReportCategory = ""
        
        
        For Each sht In wkb.Worksheets
        
            Set ReportSheet = New ReportingSheet
            
            bReportSheetAssigned = ReportSheet.AssignExistingSheet(sht)
            
            If bReportSheetAssigned And (sht.Visible = xlSheetVisible) Then
            
                'Create return to Index links
                sht.Hyperlinks.Add _
                    Anchor:=sht.Range("ReturnToIndex"), _
                    Address:="", _
                    SubAddress:="Index!A1", _
                    TextToDisplay:="<Return to Index>"
                    
                'Write the report category headers
                If ReportSheet.Category <> sLastCapturedReportCategory Then
                    i = i + 3
                    sLastCapturedReportCategory = ReportSheet.Category
                    rngCategoryCol.Cells(i) = ReportSheet.Category
                    rngCategoryCol.Cells(i).Font.Bold = True
                End If
    
                i = i + 2
                rngReportCol.Cells(i) = ReportSheet.Heading
                rngSheetNameCol.Cells(i) = sht.Name
                ActiveSheet.Hyperlinks.Add _
                    Anchor:=rngReportCol.Cells(i), _
                    Address:="", _
                    SubAddress:="'" & sht.Name & "'" & "!$F$12"
                
                'Set link to each sheets error check range (which could be empty)
                sSheetErrorCheckColumnsRangeName = "'" & sht.Name & "'!ErrorCheckColumns"
                sSheetErrorCheckRowsRangeName = "'" & sht.Name & "'!ErrorCheckRows"
                
                sIndexPageErrorCheckFormula = "=AND(" & Chr(10) & _
                 "   COUNTIFS(<ColErrCheckRange>, FALSE) = 0," & Chr(10) & _
                 "   COUNTIFS(<RowErrCheckRange>, FALSE) = 0," & Chr(10) & _
                 "   SUMPRODUCT(--ISERROR(<ColErrCheckRange>))=0," & Chr(10) & _
                 "   SUMPRODUCT(--ISERROR(<RowErrCheckRange>))=0" & Chr(10) & _
                ")"
                    
                sIndexPageErrorCheckFormula = Replace(sIndexPageErrorCheckFormula, _
                    "<ColErrCheckRange>", sSheetErrorCheckColumnsRangeName)
                sIndexPageErrorCheckFormula = Replace(sIndexPageErrorCheckFormula, _
                    "<RowErrCheckRange>", sSheetErrorCheckRowsRangeName)
                
                rngErrorCol.Cells(i).Formula = sIndexPageErrorCheckFormula

                rngErrorCol.Cells(i).Font.Color = rgb(170, 170, 170)
                
                Set ErrorCheckFormatCondition = rngErrorCol.Cells(i).FormatConditions.Add( _
                    Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE")
                With ErrorCheckFormatCondition.Font
                    .Bold = True
                    .Italic = False
                    .Color = rgb(255, 0, 0)
                    .TintAndShade = 0
                End With
                
                ReportSheet.WorkbookErrorStatusFormula = WorkbookErrorStatusFormula
                    
            End If
            
        Next sht
        
        .Range("C3").Select
        
    End With

    'Create an empty hidden first and last sheet as anchor points for 3d sum range
    'for storing sheet hashes to check completeness of index page
    Set FirstSheet = wkb.Sheets.Add(Before:=wkb.Sheets(1))
    Set LastSheet = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FirstSheet.Name = "FirstSheet"
    LastSheet.Name = "LastSheet"
    FirstSheet.Visible = xlSheetHidden
    LastSheet.Visible = xlSheetHidden
    wkb.Names.Add Name:="SumOfSheetHashes", RefersTo:="=SUM(FirstSheet:LastSheet!$D$1)"


End Sub



Private Function WorkbookErrorStatusFormula() As String

    WorkbookErrorStatusFormula = "=IFERROR(" & Chr(10) & _
    "      IF(" & Chr(10) & _
    "              COUNTIFS(Index!ErrorChecks, FALSE)=0," & Chr(10) & _
    "              ""OK""," & Chr(10) & _
    "              ""Workbook error - see index tab""" & Chr(10) & _
    "           )," & Chr(10) & _
    "      ""Error checking not set""" & Chr(10) & _
    ")"

End Function

