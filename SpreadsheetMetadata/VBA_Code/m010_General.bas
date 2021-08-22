Attribute VB_Name = "m010_General"
Option Explicit
Option Private Module
Global Const gcsMenuName As String = "SpreadsheetBI"


Sub StandardEntry()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
End Sub


Sub StandardExit()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
End Sub



Sub FormatSheet(ByRef sht As Worksheet)
'TODO  - consider removing and replacing with creation of ReportSheet object
'Applies my preferred sheet formattting

    sht.Activate

    sht.Cells.Font.Name = "Calibri"
    sht.Cells.Font.Size = 11

    sht.Range("A1").Font.Color = rgb(170, 170, 170)
    sht.Range("A1").Font.Size = 8

    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 80
    sht.DisplayPageBreaks = False
    sht.Columns("A:A").ColumnWidth = 4

    If SheetLevelRangeNameExists(sht, "SheetHeading") Then
        sht.Names("SheetHeading").Delete
    End If
    sht.Names.Add Name:="SheetHeading", RefersTo:="=$B$2"

    If SheetLevelRangeNameExists(sht, "SheetCategory") Then
        sht.Names("SheetCategory").Delete
    End If
    sht.Names.Add Name:="SheetCategory", RefersTo:="=$A$1"

    With sht.Range("SheetHeading")
        If .Value = "" Then
            .Value = "Heading"
        End If
        .Font.Bold = True
        .Font.Size = 16
    End With

End Sub


'TODO consider moving into the ReportingTable class

Sub FormatTable(lo As ListObject)

    Dim sty As TableStyle
    Dim wkb As Workbook
    
    Set wkb = lo.Parent.Parent
    
    On Error Resume Next
    wkb.TableStyles.Add ("SpreadsheetBiStyle")
    On Error GoTo 0
    Set sty = wkb.TableStyles("SpreadsheetBiStyle")
    
    'Set Header Format
    With sty.TableStyleElements(xlHeaderRow)
        .Interior.Color = rgb(68, 114, 196)
        .Font.Color = rgb(255, 255, 255)
        .Font.Bold = True
        .Borders.item(xlEdgeTop).LineStyle = xlSolid
        .Borders.item(xlEdgeTop).Weight = xlMedium
        .Borders.item(xlEdgeBottom).LineStyle = xlSolid
        .Borders.item(xlEdgeBottom).Weight = xlMedium
    End With

    'Set row stripe format
    sty.TableStyleElements(xlRowStripe1).Interior.Color = rgb(217, 217, 217)
    sty.TableStyleElements(xlRowStripe2).Interior.Color = rgb(255, 255, 255)
    
    'Set whole table bottom edge format
    sty.TableStyleElements(xlWholeTable).Borders.item(xlEdgeBottom).LineStyle = xlSolid
    sty.TableStyleElements(xlWholeTable).Borders.item(xlEdgeBottom).Weight = xlMedium

    
    'Apply custom style and set other attributes
    lo.TableStyle = "SpreadsheetBiStyle"
    With lo.HeaderRowRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
    End With
    
    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.EntireColumn.AutoFit
    End If
    
End Sub



Sub SetNumberFormat(sNumberFormat)

    If ActiveCellIsInPivotTable Then
        ActiveCell.PivotField.NumberFormat = sNumberFormat
    Else
        Selection.NumberFormat = sNumberFormat
    End If

End Sub



Function LooperValue(ByVal sItem As String) As String
'Precondition: tbl_LoopController exists in active workbook and contans column Item and Value
'This sub returns Value for corresponding sItem

    Dim sFormulaString As String
    
    sFormulaString = "=INDEX(tbl_LoopController[Value], MATCH(""" & sItem & """, tbl_LoopController[Item], 0))"
    LooperValue = Application.Evaluate(sFormulaString)


End Function



Sub SetOuterBorders(ByRef rng As Range)

    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With


End Sub




Function GetCellFontColour(rng As Range, Optional formatType As Integer = 0) As Variant

'https://stackoverflow.com/questions/24132665/return-rgb-values-from-range-interior-color-or-any-other-color-property?rq=1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function            Color
'   Purpose             Determine the Font Color Of a Cell
'   @Param rng          Range to Determine Background Color of
'   @Param formatType   Default Value = 0
'                       0   Integer
'                       1   Hex
'                       2   RGB
'                       3   Excel Color Index
'   Usage               Color(A1)      -->   9507341
'                       Color(A1, 0)   -->   9507341
'                       Color(A1, 1)   -->   91120D
'                       Color(A1, 2)   -->   13, 18, 145
'                       Color(A1, 3)   -->   6
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Dim colorVal As Variant
    colorVal = Cells(rng.Row, rng.Column).Font.Color
    Select Case formatType
        Case 1
            GetCellFontColour = Hex(colorVal)
        Case 2
            GetCellFontColour = (colorVal Mod 256) & ", " & ((colorVal \ 256) Mod 256) & ", " & (colorVal \ 65536)
        Case 3
            GetCellFontColour = Cells(rng.Row, rng.Column).Interior.ColorIndex
        Case Else
            GetCellFontColour = colorVal
    End Select
End Function





Function GetDataValidationFromRangeReference(ByVal rngSingleCell As Range) As Variant
'Returns a variant array of data validation items for rngSingleCell
'rngSingleCell must contain a list validation
    
    Dim bValidationListIsRange
    Dim rngValidationReference
    Dim sValidationFormula As String
    Dim SplitStringArray As Variant
    Dim ReturnValue() As Variant
    Dim i As Long
    
    On Error Resume Next
    sValidationFormula = rngSingleCell.Validation.Formula1
    If Err.Number <> 0 Then
        GetDataValidationFromRangeReference = Nothing
        Exit Function
    End If
    
    Set rngValidationReference = Range(Replace(sValidationFormula, "=", ""))
    bValidationListIsRange = (Err.Number = 0)
    On Error GoTo 0

    If bValidationListIsRange Then
        ReDim ReturnValue(0 To (rngValidationReference.Cells.Count - 1))
        For i = LBound(ReturnValue) To UBound(ReturnValue)
            ReturnValue(i) = rngValidationReference.Cells(i + 1)
        Next i
    Else
        SplitStringArray = Split(sValidationFormula, ",")
        ReDim ReturnValue(LBound(SplitStringArray) To UBound(SplitStringArray))
        For i = LBound(SplitStringArray) To UBound(SplitStringArray)
            ReturnValue(i) = SplitStringArray(i)
        Next i
    End If
    
    GetDataValidationFromRangeReference = ReturnValue

End Function




Sub GetPowerQueryFileNamesFromUser(ByRef FilePaths() As String)

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
        ReDim Preserve FilePaths(0 To fDialog.SelectedItems.Count - 1)
        For i = 0 To fDialog.SelectedItems.Count - 1
            FilePaths(i) = fDialog.SelectedItems(i + 1)
        Next i
    End If

End Sub


Function PowerQueryReferencedToTextFile(ByVal FileName As String) As String

    PowerQueryReferencedToTextFile = _
        "let" & vbLf & _
        "   FileName = ""<FileName>""," & vbLf & _
        "   Binary = File.Contents(FileName)," & vbLf & _
        "   QueryText = Text.FromBinary(Binary)," & vbLf & _
        "   Output = Expression.Evaluate(QueryText, #shared)" & vbLf & _
        "in" & vbLf & _
        "   Output"

    PowerQueryReferencedToTextFile = Replace(PowerQueryReferencedToTextFile, _
        "<FileName>", FileName)

End Function

