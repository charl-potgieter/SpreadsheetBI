VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReportingTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Reporting"
Option Explicit

Private Type TReportingTable
    lo As ListObject
End Type

Private this As TReportingTable
Private Const csEmptyDaxTableName As String = "EmptyTableForDaxQueries"

Public Property Get Name() As String
   Name = this.lo.Name
End Property

Public Property Let Name(sName As String)
    
    Dim Counter As Long
    Dim TryName As String
    Counter = 0
    
    If this.lo Is Nothing Then Exit Property
    
    'Get first availabe sheet name add numeric prefix if name is taken
    TryName = sName
    On Error Resume Next
    Do While this.lo.Name <> TryName
        If Counter <> 0 Then
            TryName = sName & "_" & Counter
        End If
        this.lo.Name = TryName
         Counter = Counter + 1
    Loop
    On Error GoTo 0
    
End Property


Public Sub CreateFromDaxQuery(ByVal TopLeftCell As Range, sDaxQuery As String)
    
    Dim sht As Worksheet
    Dim wkb As Workbook
    
    Set sht = TopLeftCell.Parent
    Set wkb = sht.Parent
    
    CreateEmptyPowerPivotTable wkb
    
    'If query is directly saved from DAX Studio it seems to save unicode format. Only
    'issues are symbols at start hence the replace below.  ADO can be used if more
    'issues arrise in the future.
    'https://stackoverflow.com/questions/13851473/read-utf-8-text-file-in-vbscript/13855268#13855268
    sDaxQuery = Replace(sDaxQuery, "﻿", "")

    Set this.lo = sht.ListObjects.Add( _
        SourceType:=xlSrcModel, _
        Source:=wkb.Connections(csEmptyDaxTableName), _
        Destination:=TopLeftCell.Cells(1))

    With this.lo.TableObject
        .PreserveFormatting = True
        .RefreshStyle = xlOverwriteCells
        .AdjustColumnWidth = True
    End With

    With this.lo.TableObject.WorkbookConnection.OLEDBConnection
        .CommandText = Array(sDaxQuery)
        .CommandType = xlCmdDAX
    End With

    this.lo.Refresh
    FormatReportingTable
    
End Sub


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
        .DataBodyRange.Font.Color = RGB(0, 0, 0)
    End With
    AddCalculatedField = True

End Function


Public Function FormatTableField(sFieldName As String, sFormat As String)

    On Error Resume Next
    this.lo.ListColumns(sFieldName).DataBodyRange.NumberFormat = sFormat
    FormatTableField = (Err.Number = 0)
    On Error GoTo 0

End Function


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
            ImportRelationships:=False

    End If

End Sub

