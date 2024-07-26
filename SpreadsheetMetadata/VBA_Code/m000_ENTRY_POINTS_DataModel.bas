Attribute VB_Name = "m000_ENTRY_POINTS_DataModel"
Option Explicit

Public Type TypeModelMeasures
    Name As String
    UniqueName As String
    Visible As Boolean
    Expression As String
    Table As String
End Type

Public Type TypeModelColumns
    Name As String
    UniqueName As String
    TableName As String
    Visible As Boolean
End Type

Public Type TypeModelCalcColumns
    Name As String
    TableName As String
    Expression As String
End Type

Public Type TypeModelRelationship
    ForeignKeyTable As String
    ForeignKeyColumn As String
    PrimaryKeyTable As String
    PrimaryKeyColumn As String
    Active As Boolean
End Type



Sub CreateDataModelRelationships()
'Creates relationships based on listobject in sheet Model_Relationhsips which has the below fields:
'ID, Foreign Key Table, Foreign Key Column, Primary Key Table, Primary Key Column, Active


    Dim lo As ListObject
    Dim i As Integer
    Dim mdl As Model
    Dim sForeignKeyTable As String
    Dim sForeignKeyCol As String
    Dim sPrimaryKeyTable As String
    Dim sPrimaryKeyCol As String

    StandardEntry
    Set mdl = ActiveWorkbook.Model
    Set lo = ActiveWorkbook.Sheets("Model_Relationships").ListObjects(1)
    For i = 1 To lo.DataBodyRange.Rows.Count
        sForeignKeyTable = lo.ListColumns("Foreign Key Table").DataBodyRange.Cells(i)
        sForeignKeyCol = lo.ListColumns("Foreign Key Column").DataBodyRange.Cells(i)
        sPrimaryKeyTable = lo.ListColumns("Primary Key Table").DataBodyRange.Cells(i)
        sPrimaryKeyCol = lo.ListColumns("Primary Key Table").DataBodyRange.Cells(i)
        mdl.ModelRelationships.Add _
            mdl.ModelTables(sForeignKeyTable).ModelTableColumns(sForeignKeyCol), _
            mdl.ModelTables(sPrimaryKeyTable).ModelTableColumns(sPrimaryKeyCol)
    Next i
    StandardExit
End Sub


Sub WriteModelInfoToSheets()
'Writes below information from power pivot model in activeworkbook to worksheets:
'   Model Measures
'   Model columns
'   Measure Relationships

    Dim iMsgBoxResponse As Integer

    StandardEntry

    If SheetExists(ActiveWorkbook, "ModelMeasures") Or SheetExists(ActiveWorkbook, "ModelCalcColumns") Or _
    SheetExists(ActiveWorkbook, "ModelColumns") Or SheetExists(ActiveWorkbook, "ModelRelationships") Then
            iMsgBoxResponse = MsgBox("Sheets already exists, delete?", vbQuestion + vbYesNo + vbDefaultButton2)
            If iMsgBoxResponse = vbNo Then
                Exit Sub
            End If
    End If

    WriteModelMeasuresToSheet
    WriteModelCalcColsToSheet
    WriteModelColsToSheet
    WriteModelRelationshipsToSheet

    StandardExit
End Sub

