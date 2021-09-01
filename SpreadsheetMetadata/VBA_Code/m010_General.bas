Attribute VB_Name = "m010_General"
Option Explicit
Option Private Module
Global Const gcsMenuName As String = "SpreadsheetBI"



Sub FormatSheet(ByRef sht As Worksheet)
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
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Borders.item(xlEdgeTop).LineStyle = xlSolid
        .Borders.item(xlEdgeTop).Weight = xlMedium
        .Borders.item(xlEdgeBottom).LineStyle = xlSolid
        .Borders.item(xlEdgeBottom).Weight = xlMedium
    End With

    'Set row stripe format
    sty.TableStyleElements(xlRowStripe1).Interior.Color = RGB(217, 217, 217)
    sty.TableStyleElements(xlRowStripe2).Interior.Color = RGB(255, 255, 255)
    
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




Sub ReadSavedReportSheetFormat(ByRef RptSheetFormat As TypeReportSheetFormat)

    With ThisWorkbook.Sheets("RptShtFormat")
        RptSheetFormat.SheetFont = .Range("SheetFont")
        RptSheetFormat.DefaultFontSize = .Range("DefaultFontSize")
        RptSheetFormat.ZoomPercentage = .Range("ZoomPercentage")
        RptSheetFormat.HeadingColourRed = .Range("HeadingColour_R")
        RptSheetFormat.HeadingColourGreen = .Range("HeadingColour_G")
        RptSheetFormat.HeadingColourBlue = .Range("HeadingColour_B")
        RptSheetFormat.HeadingFontSize = .Range("HeadingFontSize")
    End With

End Sub


Sub PopulateUserFormWithSavedReportSheetFormats(ByRef uf As ufReportShtFormat, _
    ByRef RptSheetFormat As TypeReportSheetFormat)

    With RptSheetFormat
        uf.tbSheetFont.Value = .SheetFont
        uf.tbDefaultfontSize = .DefaultFontSize
        uf.tbZoomPercent = .ZoomPercentage
        uf.tbHeadingColourRed = .HeadingColourRed
        uf.tbHeadingColourGreen = .HeadingColourGreen
        uf.tbHeadingColourBlue = .HeadingColourBlue
        uf.tbHeadingFontSize = .HeadingFontSize
    End With

End Sub

Sub GetReportSheetFormatFromUserForm(ByRef uf As ufReportShtFormat, _
    ByRef RptSheetFormat As TypeReportSheetFormat)

    With RptSheetFormat
        .SheetFont = uf.tbSheetFont.Value
        .DefaultFontSize = uf.tbDefaultfontSize
        .ZoomPercentage = uf.tbZoomPercent
        .HeadingColourRed = uf.tbHeadingColourRed
        .HeadingColourGreen = uf.tbHeadingColourGreen
        .HeadingColourBlue = uf.tbHeadingColourBlue
        .HeadingFontSize = uf.tbHeadingFontSize
    End With

End Sub
    




Sub WriteReportSheetFormatsToSheet(ByRef RptSheetFormat As TypeReportSheetFormat)

    With ThisWorkbook.Sheets("RptShtFormat")
        .Range("SheetFont") = RptSheetFormat.SheetFont
        .Range("DefaultFontSize") = RptSheetFormat.DefaultFontSize
        .Range("ZoomPercentage") = RptSheetFormat.ZoomPercentage
        .Range("HeadingColour_R") = RptSheetFormat.HeadingColourRed
        .Range("HeadingColour_G") = RptSheetFormat.HeadingColourGreen
        .Range("HeadingColour_B") = RptSheetFormat.HeadingColourBlue
        .Range("HeadingFontSize") = RptSheetFormat.HeadingFontSize
    End With

    ThisWorkbook.Save

End Sub


Sub ApplyReportSheetFormatProperties(ByRef ReportSht As ReportingSheet, _
    ByRef ReportSheetFormat As TypeReportSheetFormat)

    With ReportSht
        .Create ActiveWorkbook, ActiveSheet.Index
        .SheetFont = ReportSheetFormat.SheetFont
        .DefaultFontSize = ReportSheetFormat.DefaultFontSize
        .ZoomPercentage = ReportSheetFormat.ZoomPercentage
        .HeadingFontColour = Array( _
            ReportSheetFormat.HeadingColourRed, _
            ReportSheetFormat.HeadingColourGreen, _
            ReportSheetFormat.HeadingColourBlue)
        .HeadingFontSize = ReportSheetFormat.HeadingFontSize
    End With

End Sub
