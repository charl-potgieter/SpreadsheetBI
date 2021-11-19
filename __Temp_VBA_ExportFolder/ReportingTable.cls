VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReportingTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Storage.Reporting"
Option Explicit

Private Type TReportingTable
    ReportName As String
    lo As ListObject
    ReportSht As ReportingSheet
End Type

Private this As TReportingTable
Private Const csEmptyDaxTableName As String = "EmptyTableForDaxQueries"


Public Property Get Name() As String
   Name = this.ReportName
End Property


Public Property Let Name(sName As String)
    
    If this.lo Is Nothing Then Exit Property
    
    On Error Resume Next
    this.lo.Name = sName
    If Err.Number <> 0 Then MsgBox ("Reporting table name already taken, name not updated")
    On Error GoTo 0
    
End Property


Public Property Get SheetProperties() As Dictionary
    
    Dim ReturnDictionary As Dictionary
    
    Set ReturnDictionary = New Dictionary
    With ReturnDictionary
        .Add "FirstCellAddress", this.lo.HeaderRowRange.Cells(1).Address
        .Add "SheetName", this.lo.Parent.Name
        .Add "Category", this.ReportSht.Category
        .Add "Heading", this.ReportSht.Heading
        .Add "FreezePaneRow", this.ReportSht.FreezePaneRow
        .Add "FreezePaneCol", this.ReportSht.FreezePaneCol
    End With
    Set SheetProperties = ReturnDictionary
    
End Property


Public Property Let SheetProperties(ByVal Properties As Dictionary)
'Table is moved from original sheet to new sheet when sheet properties are set
    
    Dim shtOriginalSheet  As Worksheet
    Dim wkb As Workbook
    Dim sFirstCellAddress As String
    
    Set shtOriginalSheet = this.lo.Parent
    Set wkb = shtOriginalSheet.Parent
    Set this.ReportSht = New ReportingSheet
    With this.ReportSht
        .Create wkb
        .Name = Properties.item("SheetName")
        .Category = Properties.item("Category")
        .Heading = Properties.item("Heading")
        .FreezePaneRow = Properties.item("FreezePaneRow")
        .FreezePaneCol = Properties.item("FreezePaneCol")
        
        sFirstCellAddress = Properties.item("FirstCellAddress")
        this.lo.Range.Cut
        .Sheet.Paste .Sheet.Range(sFirstCellAddress)
        'if below step is ignored the this.lo reference is deleted on cut and paste
        Set this.lo = .Sheet.Range(sFirstCellAddress).ListObject
        shtOriginalSheet.Delete
    End With
    
End Property


Public Property Get Formulas() As Dictionary
    
    Dim ReturnDictionary As Dictionary
    Dim LstCol As ListColumn
    
    Set ReturnDictionary = New Dictionary
    
    For Each LstCol In this.lo.ListColumns
        If LstCol.DataBodyRange.Cells(1).HasFormula Then
            ReturnDictionary.Add LstCol.Name, LstCol.DataBodyRange.Cells(1).Formula
        End If
    Next LstCol
        
    Set Formulas = ReturnDictionary
    
End Property


Property Let Formulas(Properties As Dictionary)

    Dim Property As Variant
    Dim lc As ListColumn
    
    For Each Property In Properties.Keys
        With this.lo.ListColumns.Add
            .Name = Property
            .DataBodyRange.Formula2 = Properties(Property)
        End With
    Next Property
    
End Property


Public Property Get ColumnWidths() As Dictionary
    
    Dim ReturnDictionary As Dictionary
    Dim LstCol As ListColumn
    
    Set ReturnDictionary = New Dictionary
    
    For Each LstCol In this.lo.ListColumns
        ReturnDictionary.Add LstCol.Name, LstCol.DataBodyRange.EntireColumn.ColumnWidth
    Next LstCol
        
    Set ColumnWidths = ReturnDictionary
    
End Property


Property Let ColumnWidths(Properties As Dictionary)

    Dim Property As Variant
    Dim lc As ListColumn
    
    For Each Property In Properties.Keys
        this.lo.ListColumns(Property).DataBodyRange.EntireColumn.ColumnWidth = _
            Properties(Property)
    Next Property
    
End Property

Public Property Get NumberFormatting() As Dictionary
    
    Dim ReturnDictionary As Dictionary
    Dim LstCol As ListColumn
    
    Set ReturnDictionary = New Dictionary
    
    For Each LstCol In this.lo.ListColumns
        ReturnDictionary.Add LstCol.Name, LstCol.DataBodyRange.Cells(1).NumberFormat
    Next LstCol
        
    Set NumberFormatting = ReturnDictionary
    
End Property


Property Let NumberFormatting(Properties As Dictionary)

    Dim Property As Variant
    Dim lc As ListColumn
    
    For Each Property In Properties.Keys
        this.lo.ListColumns(Property).DataBodyRange.NumberFormat = Properties(Property)
        
    Next Property
    
End Property


Public Function AssignToExistingSheet(ByVal sht As Worksheet) As Boolean

    Dim TestReportSht As ReportingSheet
    Dim bTestReportSheetAssigned As Boolean
    
    Set TestReportSht = New ReportingSheet
    bTestReportSheetAssigned = TestReportSht.AssignExistingSheet(sht)
    
    Select Case True
        Case Not bTestReportSheetAssigned
            AssignToExistingSheet = False 'Only allow ReportingTables to exist on ReportingSheets
            Set TestReportSht = Nothing
        Case sht.ListObjects.Count <> 1
            AssignToExistingSheet = False
        Case sht.ListObjects(1).SourceType <> xlSrcModel
            AssignToExistingSheet = False
        Case Else
            AssignToExistingSheet = True
            Set this.lo = sht.ListObjects(1)
            Set this.ReportSht = New ReportingSheet
            this.ReportSht.AssignExistingSheet sht
            this.ReportName = this.ReportSht.Heading
    End Select
        
End Function



Public Sub CreateEmptyReportingTable(ByVal wkb As Workbook, ByVal sReportName As String)
    
    Dim sht As Worksheet
    Const sInitialFirstCell As String = "A1"

    CreateEmptyPowerPivotTable wkb
    Set sht = wkb.Sheets.Add
    this.ReportName = sReportName
    Set this.lo = sht.ListObjects.Add( _
        SourceType:=xlSrcModel, _
        Source:=wkb.Connections(csEmptyDaxTableName), _
        Destination:=sht.Range(sInitialFirstCell))
    With this.lo.TableObject
        .PreserveFormatting = True
        .RefreshStyle = xlOverwriteCells
        .AdjustColumnWidth = True
    End With
    FormatReportingTable
    
End Sub


Public Property Get DaxQuery() As String
    DaxQuery = this.lo.TableObject.WorkbookConnection.OLEDBConnection.CommandText
End Property


Public Property Let DaxQuery(ByVal sDaxQuery As String)

    'If query is directly saved from DAX Studio it seems to save unicode format. Only
    'issues are symbols at start hence the replace below.  ADO can be used if more
    'issues arrise in the future.
    'https://stackoverflow.com/questions/13851473/read-utf-8-text-file-in-vbscript/13855268#13855268
    sDaxQuery = Replace(sDaxQuery, "﻿", "")

    With this.lo.TableObject.WorkbookConnection.OLEDBConnection
        .CommandText = Array(sDaxQuery)
        .CommandType = xlCmdDAX
    End With

    this.lo.Refresh
    

End Property


Public Sub ConvertDaxQueryToValues()
    On Error Resume Next
    this.lo.TableObject.Delete
    On Error GoTo 0
End Sub


Public Function AddCalculatedField(ByVal sFieldName As String, ByVal sFormula As String) As Boolean
'TODO check what happens if Fieldname already exists - do I need to build in error check
'and change this into a sub that can return false if unsuccesful

    Dim listCol As ListColumn
    Dim bHeadingAlreadyExists

    bHeadingAlreadyExists = WorksheetFunction.CountIfs(this.lo.HeaderRowRange, sFieldName) <> 0
    If bHeadingAlreadyExists Then
        AddCalculatedField = False
        Exit Function
    End If
    Set listCol = this.lo.ListColumns.Add
    With listCol
        .Name = sFieldName
        .DataBodyRange.Formula = sFormula
    End With
    AddCalculatedField = True

End Function


Public Function FormatTableField(sFieldName As String, sFormat As String)

    On Error Resume Next
    this.lo.ListColumns(sFieldName).DataBodyRange.NumberFormat = sFormat
    FormatTableField = (Err.Number = 0)
    On Error GoTo 0

End Function

Public Sub ApplyColourFormatting()
'Formats DAX query text as blue, formulas as black
    
    Dim lc As ListColumn
    
    For Each lc In this.lo.ListColumns
        With lc.DataBodyRange
            If .Cells(1).HasFormula Then
                .Font.Color = RGB(0, 0, 0)
            Else
                .Font.Color = RGB(0, 112, 192)
            End If
        End With
    Next lc

End Sub


Private Sub FormatReportingTable()

    Dim sty As TableStyle
    Dim wkb As Workbook
    
    Set wkb = this.lo.Parent.Parent
    
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
    this.lo.TableStyle = "SpreadsheetBiStyle"
    With this.lo.HeaderRowRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
    End With
    
    If Not this.lo.DataBodyRange Is Nothing Then
        this.lo.DataBodyRange.EntireColumn.AutoFit
    End If
    
    'Set DAX query range font as blue (this takes place before any formula columns are added
    On Error Resume Next 'In event no data returned and then no databodyrange
    this.lo.DataBodyRange.Font.Color = RGB(0, 112, 192)
    On Error GoTo 0
    
End Sub



Private Sub CreateEmptyPowerPivotTable(ByVal wkb As Workbook)
'Creates empty table loaded into Power Pivot to form source of listobject connection for DAX queries

    Dim sQueryStr As String
    Dim sTestName As String
    Dim bEmptyTableExists As Boolean

    On Error Resume Next
        sTestName = wkb.Model.ModelTables(csEmptyDaxTableName).Name
        If Err.Number = 0 Then
            bEmptyTableExists = True
        Else
            bEmptyTableExists = False
            wkb.Queries(csEmptyDaxTableName).Delete
        End If
    On Error GoTo 0


    If Not bEmptyTableExists Then
        sQueryStr = "let" & vbCr & _
            "    tbl = Table.FromRecords({[NullHeader = """"]}), " & vbCr & _
            "    ChangedType = Table.TransformColumnTypes(tbl, {{""NullHeader"", type text}})" & vbCr & _
            "in" & vbCr & _
            "    ChangedType"
    
        wkb.Queries.Add Name:=csEmptyDaxTableName, Formula:=sQueryStr
        wkb.Connections.Add2 _
            Name:=csEmptyDaxTableName, _
            Description:="Connection to the 'EmptyTable' query in the workbook.", _
            ConnectionString:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & _
                csEmptyDaxTableName & ";Extended Properties=", _
            CommandText:="" & csEmptyDaxTableName & "", _
            lCmdtype:=6, _
            CreateModelConnection:=True, _
            ImportRelationShips:=False

    End If

End Sub


