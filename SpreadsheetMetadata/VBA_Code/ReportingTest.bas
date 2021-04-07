Attribute VB_Name = "ReportingTest"
'@Folder "Reporting"
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
    bSheetImported = NewFormattedSheet.CreateFromExistingSheet(ActiveWorkbook.Sheets.item(1))
    
    
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



Public Sub TestTableDax()

    Dim rt As ReportingTable
    Set rt = New ReportingTable
    Dim bFieldAdded As Boolean
    
    rt.CreateFromDaxQuery ActiveCell, "EVALUATE SUMMARIZECOLUMNS(DimAccounts[Class],DimAccounts[Group],""Actual Amt"", [Actual Amt])"
    rt.ConvertDaxQueryToValues
    bFieldAdded = rt.AddCalculatedField("Double", "=[@[Actual Amt]] * 2")
    rt.FormatTableField "Double", "#,##0_);(#,##0);-??"

End Sub




Sub TestPP2()

'    Dim pp As PivotReport
'    Dim CubeField As CubeField
'    Dim dict As Dictionary
'    Dim key As Variant
'
'    Set pp = New PivotReport
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

'    Dim pp As PivotReport
'    Dim CubeFieldName As Variant
'    Dim cf As CubeField
'
'    Set pp = New PivotReport
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

