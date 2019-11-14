Attribute VB_Name = "m000_EntryPoints"
Option Explicit


Sub FormatZeroDecimalNumberFormat()
    SetNumberFormat "#,##0_);(#,##0);-??"
End Sub



Sub FormatOneDecimalNumberFormat()
    SetNumberFormat "#,##0.0_);(#,##0.0);-??"
End Sub



Sub FormatTwoDecimalsNumberFormat()
    SetNumberFormat "#,##0.00_);(#,##0.00);-??"
End Sub



Sub FormatTwoDigitPercentge()
    SetNumberFormat "0.00%"
End Sub


Sub FormatFourDigitPercentge()
    SetNumberFormat "0.0000%"
End Sub


Sub FormatDate()
    SetNumberFormat "dd-mmm-yy"
End Sub



Sub InsertFormattedSheetIntoActiveWorkbook()
    
    Dim sht As Worksheet
    
    Set sht = ActiveWorkbook.Sheets.Add(After:=ActiveSheet)
    FormatSheet sht

End Sub


Sub FormatActiveSheet()

    FormatSheet ActiveSheet

End Sub



Sub FormatHeadings()

    Application.ScreenUpdating = False
    
    'Remove all current borders
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    'Set new borders
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    'Set header colour
    With Selection.Interior
        .Color = RGB(255, 255, 153)
        .Pattern = xlSolid
    End With
    
    Selection.Font.Bold = True
    
    'Set Text allignment
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    
    
    Application.ScreenUpdating = True


End Sub



Sub ExportPowerQueriesInActiveWorkbookToFiles()

    Dim sFolderSelected As String
    
    sFolderSelected = GetFolder
    If NumberOfFilesInFolder(sFolderSelected) <> 0 Then
        MsgBox ("Please select an empty folder...exiting")
        Exit Sub
    End If
    
    ExportPowerQueriesToFiles sFolderSelected, ActiveWorkbook

End Sub




Sub ImportPowerQueriesFromSelectedFolderNonRecursive()

    Dim sFolderSelected As String
    
    sFolderSelected = GetFolder
    ImportOrRefreshPowerQueriesInFolder sFolderSelected, False

End Sub



Sub ImportPowerQueriesFromSelectedFolderRecursive()

    Dim sFolderSelected As String
    
    sFolderSelected = GetFolder
    ImportOrRefreshPowerQueriesInFolder sFolderSelected, True

End Sub



Sub ImportSelectedPowerQueries()
'Requires reference to Microsoft Scripting runtime library

    Dim sPowerQueryFilePath As String
    Dim sPowerQueryName As String
    Dim fDialog As FileDialog
    Dim fso As FileSystemObject
    Dim i As Integer
    
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With fDialog
        .AllowMultiSelect = True
        .Title = "Select power query / queries"
        .InitialFileName = ThisWorkbook.Path
        .Filters.Clear
        .Filters.Add "m Power Query Files", "*.m"
    End With
    
    'fDialog.Show value of -1 below means success
    If fDialog.Show = -1 Then
        For i = 1 To fDialog.SelectedItems.Count
            sPowerQueryFilePath = fDialog.SelectedItems(i)
            Set fso = New FileSystemObject
            sPowerQueryName = Replace(fso.GetFileName(sPowerQueryFilePath), ".m", "")
            ImportOrRefreshSinglePowerQuery sPowerQueryFilePath, sPowerQueryName, ActiveWorkbook
        Next i
    End If
    

End Sub




Sub CreateIndexPage()

    Dim sht As Worksheet
    Dim shtIndex As Worksheet
    Dim i As Double
    Dim sPreviousReportCategory As String
    Dim sReportCategory As String
    Dim sReportName As String
    Dim rngCategoryCol As Range
    Dim rngReportCol As Range
    Dim rngSheetNameCol As Range
    Dim rngShowRange As Range
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    'Delete any previous index sheet and create a new one
    On Error Resume Next
    ActiveWorkbook.Sheets("Index").Delete
    On Error GoTo 0
    Set shtIndex = ActiveWorkbook.Sheets.Add(Before:=ActiveWorkbook.Sheets(1))
    FormatSheet shtIndex
    
    With shtIndex
    
        .Name = "Index"
        .Range("A:A").Insert Shift:=xlToRight
        .Range("A:A").EntireColumn.Hidden = True
        .Range("C2") = "Index"
        .Range("D5").Font.Bold = True
        .Columns("D:D").ColumnWidth = 100
        .Rows("4:4").Select
        ActiveWindow.FreezePanes = True
        
        Set rngSheetNameCol = .Columns("A")
        Set rngCategoryCol = .Columns("C")
        Set rngReportCol = .Columns("D")
       
        sPreviousReportCategory = ""
        i = 2
        
        
        For Each sht In ActiveWorkbook.Worksheets
        
            sReportCategory = sht.Range("A1")
            sReportName = sht.Range("B2")
            
            If (sReportCategory <> "" And sReportName <> "") And (sht.Name <> "Index") And (sht.Visible = xlSheetVisible) Then
            
                'Create return to Index links
                sht.Hyperlinks.Add _
                    Anchor:=sht.Range("B3"), _
                    Address:="", _
                    SubAddress:="Index!A1", _
                    TextToDisplay:="<Return to Index>"
                    
                'Write the report category headers
                If sReportCategory <> sPreviousReportCategory Then
                    i = i + 3
                    rngCategoryCol.Cells(i) = sReportCategory
                    rngCategoryCol.Cells(i).Font.Bold = True
                    sPreviousReportCategory = sReportCategory
                End If
    
                i = i + 2
                rngReportCol.Cells(i) = sReportName
                rngSheetNameCol.Cells(i) = sht.Name
                
                ActiveSheet.Hyperlinks.Add _
                    Anchor:=rngReportCol.Cells(i), _
                    Address:="", _
                    SubAddress:="'" & sht.Name & "'" & "!B$4"
                    
            End If
            
        Next sht
        
        .Range("C3").Select
        
    End With
        
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

End Sub














