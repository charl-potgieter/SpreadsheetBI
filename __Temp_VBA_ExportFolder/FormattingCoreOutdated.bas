Attribute VB_Name = "FormattingCoreOutdated"
'@Folder "Formatting"
Option Explicit
'Option Private Module

    

Public Sub FormatSheet(ByRef sht As Worksheet)
'TODO  - consider removing and replacing with creation of ReportSheet object
'Applies my preferred sheet formattting

    sht.Activate

    sht.Cells.Font.Name = "Calibri"
    sht.Cells.Font.Size = 11

    sht.Range("A1").Font.Color = RGB(170, 170, 170)
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



Public Sub FormatTable(lo As ListObject)

    Dim sty As TableStyle
    Dim wkb As Workbook
    
    Set wkb = lo.Parent.Parent
    
    On Error Resume Next
    wkb.TableStyles.Add ("SpreadsheetBiStyle")
    On Error GoTo 0
    Set sty = wkb.TableStyles("SpreadsheetBiStyle")
    
    'Set Header Format
    With sty.TableStyleElements(xlHeaderRow)
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Borders.Item(xlEdgeTop).LineStyle = xlSolid
        .Borders.Item(xlEdgeTop).Weight = xlMedium
        .Borders.Item(xlEdgeBottom).LineStyle = xlSolid
        .Borders.Item(xlEdgeBottom).Weight = xlMedium
    End With

    'Set row stripe format
    sty.TableStyleElements(xlRowStripe1).Interior.Color = RGB(217, 217, 217)
    sty.TableStyleElements(xlRowStripe2).Interior.Color = RGB(255, 255, 255)
    
    'Set whole table bottom edge format
    sty.TableStyleElements(xlWholeTable).Borders.Item(xlEdgeBottom).LineStyle = xlSolid
    sty.TableStyleElements(xlWholeTable).Borders.Item(xlEdgeBottom).Weight = xlMedium

    
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


'
'Public Sub SetNumberFormat(sNumberFormat)
'
'    If ActiveCellIsInPivotTable Then
'        ActiveCell.PivotField.NumberFormat = sNumberFormat
'    Else
'        Selection.NumberFormat = sNumberFormat
'    End If
'
'    ActiveCell.PivotField.LabelRange.Addr
'
'End Sub

