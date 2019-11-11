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
        If .Value = "" Then
            .Value = "Heading"
        End If
        .Font.Bold = True
        .Font.Size = 16
    End With

End Sub


Function SheetLevelRangeNameExists(sht As Worksheet, ByRef sRangeName As String)
'Returns TRUE if sheet level scoped range name exists

    Dim sTest As String
    
    On Error Resume Next
    sTest = sht.Names(sRangeName).Name
    SheetLevelRangeNameExists = (Err.Number = 0)
    On Error GoTo 0


End Function



Sub CreateDaxQueryTable()

    Dim lo As ListObject
    
    Set lo = ActiveSheet.ListObjects.Add( _
        SourceType:=xlSrcModel, _
        Source:=ActiveWorkbook.Connections("Query - CreatedTable"), _
        Destination:=Range("$B$5"))
    
    lo.TableObject.WorkbookConnection.OLEDBConnection.CommandType = xlCmdDAX
    lo.TableObject.WorkbookConnection.OLEDBConnection.CommandText = "EVALUATE VALUES(CreatedTable[Text Column])"
    
    lo.TableObject.Refresh
    

End Sub

