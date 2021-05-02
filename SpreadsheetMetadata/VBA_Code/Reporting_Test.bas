Attribute VB_Name = "Reporting_Test"
'@Folder "Storage.Reporting"
'@IgnoreModule
'TODO remove this module
Option Explicit

Sub TestSheet1()

    Dim Report As ReportingSheet
    
    Set Report = New ReportingSheet
    Report.Create ActiveWorkbook


End Sub


Public Sub TestSheet2()

    Dim NewFormattedSheet As ReportingSheet
    Dim SecondReportingSheet As ReportingSheet
    Dim bSheetImported As Boolean
    
    Set NewFormattedSheet = New ReportingSheet
    NewFormattedSheet.Create ActiveWorkbook, 1
    
    NewFormattedSheet.Name = "SheetName"
    NewFormattedSheet.Heading = "MyHeading"
    NewFormattedSheet.Category = "MyCat"
    NewFormattedSheet.FreezePaneRow = 4
    NewFormattedSheet.FreezePaneCol = 3
    
    Set SecondReportingSheet = New ReportingSheet
    bSheetImported = NewFormattedSheet.CreateFromExistingSheet(ActiveWorkbook.Sheets.Item(1))
    
    
End Sub


Public Sub TestTable()

'    Dim tbl As New ReportingTable
'    Set tbl = New ReportingTable
'
'    tbl.Create Selection
End Sub




Public Sub TestSheet3()

    Dim NewFormattedSheet As ReportingSheet
    Dim bSheetImported As Boolean
    Dim rng As Range
    
    Set NewFormattedSheet = New ReportingSheet
    NewFormattedSheet.Create ActiveWorkbook, 1
    
    NewFormattedSheet.Name = "SheetName"
    NewFormattedSheet.Heading = "MyHeading"
    NewFormattedSheet.Category = "MyCat"
    NewFormattedSheet.FreezePaneRow = 4
    NewFormattedSheet.FreezePaneCol = 3
    
    
    
End Sub







Sub TestPP2()

'    Dim pp As ReportingPowerPivot
'    Dim CubeField As CubeField
'    Dim dict As Dictionary
'    Dim key As Variant
'
'    Set pp = New ReportingPowerPivot
'    pp.AssignToExistingPivot ActiveSheet.PivotTables(1)
'    For Each CubeField In pp.CubeFields
'        If CubeField.Orientation <> xlHidden Then
'            Debug.Print CubeField.Name
'            Debug.Print "-------------------------"
'            Set dict = pp.CubeFieldProperties(CubeField)
'            For Each key In dict
'                Debug.Print key & " | " & dict(key)
'            Next key
'            Debug.Print vbCrLf
'        End If
'    Next CubeField


End Sub


Sub TestPP3()

'    Dim pp As ReportingPowerPivot
'    Dim CubeFieldName As Variant
'    Dim cf As CubeField
'
'    Set pp = New ReportingPowerPivot
'    pp.AssignToExistingPivot ActiveSheet.PivotTables(1)
'    For Each cf In pp.CubeFields
'        If cf.Orientation <> xlHidden Then
'            Debug.Print cf.Name
'        End If
'    Next cf


End Sub





Sub test()

    Dim pvt As PivotTable

    Set pvt = ActiveCell.PivotTable
    
    pvt.TableRange2.Cut
    ActiveSheet.Paste ActiveCell

End Sub



Sub TestConInModel()

    Dim con As WorkbookConnection
    
    'TODO - need to check if
    For Each con In ActiveWorkbook.Connections
        If con.Type = xlConnectionTypeOLEDB Then
            Debug.Print con.OLEDBConnection.CommandText & " : " & con.InModel
        End If
    Next con

End Sub


Sub CheckForDaxQueryTable()

    Dim lo As ListObject
    
    Set lo = ActiveSheet.ListObjects(1)
    Debug.Print (lo.SourceType = xlSrcModel)
    

'    Set lo = sht.ListObjects.Add( _
'        SourceType:=xlSrcModel, _
'        Source:=wkb.Connections(csEmptyDaxTableName), _
'        Destination:=TopLeftCell.Cells(1))
'
'    With this.lo.TableObject
'        .PreserveFormatting = True
'        .RefreshStyle = xlOverwriteCells
'        .AdjustColumnWidth = True
'    End With
'
'    With this.lo.TableObject.WorkbookConnection.OLEDBConnection
'        .CommandText = Array(sDaxQuery)
'        .CommandType = xlCmdDAX
'    End With


End Sub



Public Sub TestTableDax()

    Dim rt As ReportingTable
    Set rt = New ReportingTable
    Dim bFieldAdded As Boolean
    
    rt.CreateEmptyReportingTable ActiveWorkbook, "test"
    rt.DaxQuery = "EVALUATE SUMMARIZECOLUMNS(DimAccounts[Class],DimAccounts[Group],""Actual Amt"", [Actual Amt])"
    Debug.Print (rt.DaxQuery)
    
    
    'rt.Create ActiveCell, "EVALUATE SUMMARIZECOLUMNS(DimAccounts[Class],DimAccounts[Group],""Actual Amt"", [Actual Amt])"
    'rt.ConvertDaxQueryToValues
'    bFieldAdded = rt.AddCalculatedField("Double", "=[@[Actual Amt]] * 2")
'    rt.FormatTableField "Double", "#,##0_);(#,##0);-??"
'    rt.Name = "MyTable"

End Sub


Public Sub TestReportingTable()
    
    Dim d As Dictionary
    Dim rt As ReportingTable

    Set rt = New ReportingTable
    rt.AssignToExistingSheet ActiveSheet
    'Set d = New Dictionary
    Set d = rt.NumberFormatting

End Sub
