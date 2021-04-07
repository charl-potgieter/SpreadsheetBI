Attribute VB_Name = "Tests"
'@Folder "Reporting"
'@IgnoreModule
Option Explicit


Public Sub TestSheet()

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
    NewFormattedSheet.Format
    
    Set SecondReportingSheet = New ReportingSheet
    bSheetImported = NewFormattedSheet.ImportExistingSheet(ActiveWorkbook.Sheets.Item(1))
    
    
End Sub


Public Sub TestTable()

    Dim tbl As New ReportingTable
    Set tbl = New ReportingTable
    
    tbl.Create Selection
End Sub




Public Sub TestTable2()

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
    NewFormattedSheet.Format
    
    
    Set rng = NewFormattedSheet.Sheet.Range("B5")
    NewFormattedSheet.AddReportingTable rng
    
    
End Sub
