Attribute VB_Name = "m010_General"
Option Explicit

Sub FormatSheet(ByRef sht As Worksheet)
'Applies my preferred sheet formattting

    sht.Activate
    
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
    
    With sht.Range("SheetHeading")
        If .value = "" Then
            .value = "Heading"
        End If
        .Font.Bold = True
        .Font.Size = 16
    End With

End Sub




