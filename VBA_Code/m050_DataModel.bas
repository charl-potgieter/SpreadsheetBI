Attribute VB_Name = "m050_DataModel"
Option Explicit
Option Private Module

Sub ExportPowerQueriesToFiles(ByVal sFolderPath As String, wkb As Workbook)

    Dim qry As WorkbookQuery
    
    For Each qry In wkb.Queries
        WriteStringToTextFile qry.Formula, sFolderPath & Application.PathSeparator & qry.Name & ".m"
    Next qry

End Sub

Sub ExportNonStandardPowerQueriesToFiles(ByVal sFolderPath As String, wkb As Workbook)
'Exports power queries without fn_std or template_std prefix

    Dim qry As WorkbookQuery
    
    
    For Each qry In wkb.Queries
        If (InStr(1, UCase(qry.Name), "FN_STD_") = 0) And (InStr(1, UCase(qry.Name), "TEMPLATE_STD_") = 0) Then
            WriteStringToTextFile qry.Formula, sFolderPath & Application.PathSeparator & qry.Name & ".m"
        End If
    Next qry

End Sub


Sub ImportOrRefreshSinglePowerQuery(ByVal sQueryPath As String, ByVal sQueryName As String, Optional wkb As Workbook)

    Dim sQueryText As String
    
    If wkb Is Nothing Then Set wkb = ActiveWorkbook
    
    sQueryText = ReadTextFileIntoString(sQueryPath)
    If QueryExists(sQueryName) Then
        wkb.Queries(sQueryName).Formula = sQueryText
    Else
        wkb.Queries.Add sQueryName, sQueryText
    End If
        

End Sub

Sub ImportOrRefreshPowerQueriesInFolder(ByVal sFolderPath As String, ByVal bRecursive As Boolean)
'Reference: Microsoft Scripting Runtime
    
    Dim FileItems() As Scripting.File
    Dim FileItem
    Dim sQueryName As String
    
    FileItemsInFolder sFolderPath, bRecursive, FileItems
    
    For Each FileItem In FileItems
        sQueryName = Left(FileItem.Name, Len(FileItem.Name) - 2)
        ImportOrRefreshSinglePowerQuery FileItem.Path, sQueryName, ActiveWorkbook
    Next FileItem


End Sub



Function TableExistInDataModel(wkb As Workbook, ByVal sTableName As String) As Boolean

    Dim sTableNameTest As String
    
    On Error Resume Next
    sTableName = wkb.Model.ModelTables(sTableName).Name
    TableExistInDataModel = (Err.Number = 0)
    On Error GoTo 0
    

End Function


Sub LoadQueryToDataModel(ByVal sQueryName As String, Optional wkb As Workbook)

    If wkb Is Nothing Then Set wkb = ActiveWorkbook
    
    If Not TableExistInDataModel(wkb, sQueryName) Then
        ActiveWorkbook.Connections.Add2 _
            Name:=sQueryName, _
            Description:=sQueryName, _
            ConnectionString:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & sQueryName & ";Extended Properties=", _
            CommandText:=sQueryName, _
            lCmdType:=6, _
            CreateModelConnection:=True, _
            ImportRelationships:=False
    End If
            
End Sub



Sub AddMeasureToDataModel(ByVal sMeasureName As String, ByVal sTableName As String, _
    ByVal sMeasure As String, ByVal sDescription As String, Optional ByRef wkb As Workbook)
    
        
    If wkb Is Nothing Then Set wkb = ActiveWorkbook
    
    wkb.Model.ModelMeasures.Add _
        MeasureName:=sMeasureName, _
        AssociatedTable:=wkb.Model.ModelTables(sTableName), _
        Formula:=1, _
        FormatInformation:=wkb.Model.ModelFormatGeneral, _
        Description:=sDescription
    
End Sub




Sub CreateDaxQueryTable()

    Dim lo As ListObject
    
    'Set source as first connection in the data model.  Seems like a connection needs to be assigned but
    'is irrelevant which table as query is determined by the DAX string anyway
    Set lo = ActiveSheet.ListObjects.Add( _
        SourceType:=xlSrcModel, _
        Source:=ActiveWorkbook.Connections(ActiveWorkbook.Model.ModelTables(1).SourceWorkbookConnection.Name), _
        Destination:=Range("$B$5"))
    
    lo.TableObject.WorkbookConnection.OLEDBConnection.CommandType = xlCmdDAX
    lo.TableObject.WorkbookConnection.OLEDBConnection.CommandText = "EVALUATE VALUES(CreatedTable[Text Column])"
    
    lo.TableObject.Refresh
    
End Sub


Sub CopyPowerQueriesBetweenFiles(ByRef wkbSource As Workbook, ByRef wkbTarget As Workbook)


    Dim qry As WorkbookQuery
    
    For Each qry In wkbSource.Queries
        If QueryExists(qry.Name, wkbTarget) Then
            wkbTarget.Queries(qry.Name).Formula = qry.Formula
        Else
            wkbTarget.Queries.Add qry.Name, qry.Formula
        End If
    Next qry


End Sub


Sub WritesMeasuresColumnsRelationshipsToSheets()

    Dim aMeasures() As TypeModelMeasures
    Dim aColumns() As TypeModelColumns
    Dim aCalcColumns() As TypeModelCalcColumns
    Dim aModelRelationships() As TypeModelRelationship
    Dim rngValidations As Range
    Dim lo As ListObject
    Dim i As Integer
    Dim j As Integer

      
    '----------------- Populate Model Measures Sheet and write visible measures to validation sheet ---------------
    
    GetModelMeasures aMeasures
    Set lo = ActiveWorkbook.Sheets("ModelMeasures").ListObjects("tbl_ModelMeasures")
    lo.DataBodyRange.ClearContents
    lo.DataBodyRange.Offset(1, 0).EntireRow.Delete
    Set rngValidations = ActiveWorkbook.Sheets("Validations").Range("val_Measures")
    rngValidations.ClearContents
    
    j = 0
    With lo
        For i = 0 To UBound(aMeasures)
            .ListColumns("Name").DataBodyRange.Cells(i + 1) = aMeasures(i).Name
            .ListColumns("Visible").DataBodyRange.Cells(i + 1) = aMeasures(i).Visible
            .ListColumns("Unique Name").DataBodyRange.Cells(i + 1) = aMeasures(i).UniqueName
            .ListColumns("DAX Expression").DataBodyRange.Cells(i + 1) = "':=" & aMeasures(i).Expression
            .ListColumns("Name and Expression").DataBodyRange.Cells(i + 1) = aMeasures(i).Name & ":=" & aMeasures(i).Expression
            If aMeasures(i).Visible Then
                rngValidations.Cells(1).Offset(j) = aMeasures(i).UniqueName
                j = j + 1
            End If
        Next i
    End With
    
    ActiveWorkbook.Names("val_Measures").RefersTo = "=Validations!" & rngValidations.Cells(1).Resize(j).Address
    

    '----------------- Populate Model Columns Sheet and write visible columns to validation sheet ---------------

    GetModelColumns aColumns
    Set lo = ActiveWorkbook.Sheets("ModelColumns").ListObjects("tbl_ModelColumns")
    lo.DataBodyRange.ClearContents
    lo.DataBodyRange.Offset(1, 0).EntireRow.Delete
    Set rngValidations = ActiveWorkbook.Names("val_Columns").RefersToRange
    rngValidations.ClearContents

    
    j = 0
    With lo
        For i = 0 To UBound(aColumns)
            .ListColumns("Name").DataBodyRange.Cells(i + 1) = aColumns(i).Name
            .ListColumns("Table Name").DataBodyRange.Cells(i + 1) = aColumns(i).TableName
            .ListColumns("Unique Name").DataBodyRange.Cells(i + 1) = aColumns(i).UniqueName
            .ListColumns("Visible").DataBodyRange.Cells(i + 1) = aColumns(i).Visible
            .ListColumns("Is calculated column").DataBodyRange.Cells(i + 1).Formula = "=COUNTIFS(tbl_ModelCalcColumns[Name], [@Name]) = 1"
            If aColumns(i).Visible Then
                rngValidations.Cells(1).Offset(j) = aColumns(i).UniqueName
                j = j + 1
            End If
        Next i
    End With
    
    ActiveWorkbook.Names("val_Columns").RefersTo = "=Validations!" & rngValidations.Cells(1).Resize(j).Address
    
    
    '----------------- Populate Model Calculated Columns Sheet ---------------

    GetModelCalculatedColumns aCalcColumns
    Set lo = ActiveWorkbook.Sheets("ModelCalcColumns").ListObjects("tbl_ModelCalcColumns")
    lo.DataBodyRange.ClearContents
    lo.DataBodyRange.Offset(1, 0).EntireRow.Delete
    
    With lo
        For i = 0 To UBound(aCalcColumns)
            .ListColumns("Name").DataBodyRange.Cells(i + 1) = aCalcColumns(i).Name
            .ListColumns("Table Name").DataBodyRange.Cells(i + 1) = aCalcColumns(i).TableName
            .ListColumns("Expression").DataBodyRange.Cells(i + 1) = aCalcColumns(i).Expression
        Next i
    End With
    
    
    '----------------- Populate Model Relationship Sheet ---------------

    GetModelRelationships aModelRelationships
    Set lo = ActiveWorkbook.Sheets("ModelRelationships").ListObjects("tbl_ModelRelationships")
    lo.DataBodyRange.ClearContents
    lo.DataBodyRange.Offset(1, 0).EntireRow.Delete
    
    With lo
        For i = 0 To UBound(aModelRelationships)
            .ListColumns("Primary Key Table").DataBodyRange.Cells(i + 1) = aModelRelationships(i).PrimaryKeyTable
            .ListColumns("Primary Key Column").DataBodyRange.Cells(i + 1) = aModelRelationships(i).PrimaryKeyColumn
            .ListColumns("Foreign Key Table").DataBodyRange.Cells(i + 1) = aModelRelationships(i).ForeignKeyTable
            .ListColumns("Foreign Key Column").DataBodyRange.Cells(i + 1) = aModelRelationships(i).ForeignKeyColumn
            .ListColumns("Active").DataBodyRange.Cells(i + 1) = aModelRelationships(i).Active
        Next i
    End With
    

End Sub




Sub GetModelColumnNames(ByRef asColumnList() As String, Optional bReturnVisibleOnly As Boolean = True)
'Requires reference to Microsoft ActiveX Data Objects
'Returns columns in the data model

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sht As Excel.Worksheet
    Dim iRowNum As Integer
    Dim i As Integer
    Dim sSQL As String

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    i = 0

    ' SQL like query to get result of DMV from schema $SYSTEM
    If bReturnVisibleOnly Then
        sSQL = "select [HIERARCHY_UNIQUE_NAME] from $SYSTEM.MDSCHEMA_HIERARCHIES " & _
            "WHERE [HIERARCHY_UNIQUE_NAME] <> '[MEASURES]' AND " & _
            "[CUBE_NAME] = 'MODEL' AND " & _
            "[HIERARCHY_IS_VISIBLE] " & _
            "ORDER BY [HIERARCHY_UNIQUE_NAME]"
    Else
        sSQL = "select [HIERARCHY_UNIQUE_NAME] from $SYSTEM.MDSCHEMA_HIERARCHIES " & _
            "WHERE [HIERARCHY_UNIQUE_NAME] <> '[MEASURES]' AND " & _
            "[CUBE_NAME] = 'MODEL' " & _
            "ORDER BY [HIERARCHY_UNIQUE_NAME]"
    End If

    ' Open connection to PowerPivot engine
    Set conn = ActiveWorkbook.Model.DataModelConnection.ModelConnection.ADOConnection
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.Open sSQL, conn, adOpenForwardOnly, adLockOptimistic
        
    If rs.RecordCount > 0 Then
        ReDim asColumnList(0 To rs.RecordCount - 1)
        ' Output of the query results
        Do Until rs.EOF
            asColumnList(i) = rs.Fields(0).Value
            rs.MoveNext
            i = i + 1
        Loop
    End If

    
    rs.Close
    Set rs = Nothing
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True


End Sub


Sub GetModelMeasureNames(ByRef asMeasureList() As String, Optional bReturnVisibleOnly As Boolean = True)
'Requires reference to Microsoft ActiveX Data Objects
'Returns measures names in the data model

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sht As Excel.Worksheet
    Dim iRowNum As Integer
    Dim i As Integer
    Dim sSQL As String

    i = 0

    ' SQL like query to get result of DMV from schema $SYSTEM
    If bReturnVisibleOnly Then
        sSQL = "select [MEASURE_UNIQUE_NAME] from $SYSTEM.MDSCHEMA_MEASURES  " & _
            "WHERE LEN([EXPRESSION]) > 0 AND " & _
            "[EXPRESSION] <> '1' AND " & _
            "[MEASURE_IS_VISIBLE] " & _
            "ORDER BY [MEASURE_UNIQUE_NAME]"
    Else
        sSQL = "select [MEASURE_UNIQUE_NAME] from $SYSTEM.MDSCHEMA_MEASURES  " & _
            "WHERE LEN([EXPRESSION]) > 0 AND " & _
            "[EXPRESSION] <> '1' " & _
            "ORDER BY [MEASURE_UNIQUE_NAME]"
    End If

    ' Open connection to PowerPivot engine
    Set conn = ActiveWorkbook.Model.DataModelConnection.ModelConnection.ADOConnection
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.Open sSQL, conn, adOpenForwardOnly, adLockOptimistic
        
    If rs.RecordCount > 0 Then
        ReDim asMeasureList(0 To rs.RecordCount - 1)
        ' Output of the query results
        Do Until rs.EOF
            asMeasureList(i) = rs.Fields(0).Value
            rs.MoveNext
            i = i + 1
        Loop
    End If
    
    rs.Close
    Set rs = Nothing


End Sub


Sub GetModelMeasures(ByRef aModelMeasures() As TypeModelMeasures)
'Requires reference to Microsoft ActiveX Data Objects
'Returns measures in the data model

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sht As Excel.Worksheet
    Dim iRowNum As Integer
    Dim i As Integer
    Dim sSQL As String

    i = 0

    ' SQL like query to get result of DMV from schema $SYSTEM
     sSQL = "select [MEASURE_NAME], [MEASURE_UNIQUE_NAME], [MEASURE_IS_VISIBLE], [EXPRESSION] from $SYSTEM.MDSCHEMA_MEASURES  " & _
            "WHERE LEN([EXPRESSION]) > 0 AND " & _
            "[MEASURE_NAME] <> '__No measures defined' " & _
            "ORDER BY [MEASURE_UNIQUE_NAME]"
    

    ' Open connection to PowerPivot engine
    Set conn = ActiveWorkbook.Model.DataModelConnection.ModelConnection.ADOConnection
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.Open sSQL, conn, adOpenForwardOnly, adLockOptimistic
        
    If rs.RecordCount > 0 Then
        ReDim aModelMeasures(0 To rs.RecordCount - 1)
        ' Output of the query results
        Do Until rs.EOF
            aModelMeasures(i).Name = rs.Fields("MEASURE_NAME")
            aModelMeasures(i).Visible = rs.Fields("MEASURE_IS_VISIBLE")
            aModelMeasures(i).UniqueName = rs.Fields("MEASURE_UNIQUE_NAME")
            aModelMeasures(i).Expression = rs.Fields("EXPRESSION")
            rs.MoveNext
            i = i + 1
        Loop
    End If
    
    rs.Close
    Set rs = Nothing


End Sub


Sub GetModelColumns(ByRef aModelColumns() As TypeModelColumns)
'Requires reference to Microsoft ActiveX Data Objects
'Returns columns the data model

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sht As Excel.Worksheet
    Dim iRowNum As Integer
    Dim i As Integer
    Dim sSQL As String

    i = 0

    ' SQL like query to get result of DMV from schema $SYSTEM
    sSQL = "select [DIMENSION_UNIQUE_NAME], [HIERARCHY_NAME], [HIERARCHY_UNIQUE_NAME], [DIMENSION_IS_VISIBLE] from $SYSTEM.MDSCHEMA_HIERARCHIES " & _
        "WHERE [HIERARCHY_UNIQUE_NAME] <> '[MEASURES]' AND " & _
        "[CUBE_NAME] = 'MODEL' " & _
        "ORDER BY [HIERARCHY_UNIQUE_NAME]"

    ' Open connection to PowerPivot engine
    Set conn = ActiveWorkbook.Model.DataModelConnection.ModelConnection.ADOConnection
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.Open sSQL, conn, adOpenForwardOnly, adLockOptimistic
        
    If rs.RecordCount > 0 Then
        ReDim aModelColumns(0 To rs.RecordCount - 1)
        ' Output of the query results
        Do Until rs.EOF
            aModelColumns(i).Name = rs.Fields("HIERARCHY_NAME")
            aModelColumns(i).UniqueName = rs.Fields("HIERARCHY_UNIQUE_NAME")
            aModelColumns(i).Visible = rs.Fields("DIMENSION_IS_VISIBLE")
            aModelColumns(i).TableName = Replace(Replace(rs.Fields("DIMENSION_UNIQUE_NAME"), "[", ""), "]", "")
            rs.MoveNext
            i = i + 1
        Loop
    End If
    
    rs.Close
    Set rs = Nothing


End Sub

Sub GetModelCalculatedColumns(ByRef aCalcColumns() As TypeModelCalcColumns)
'Requires reference to Microsoft ActiveX Data Objects
'Returns calcualted columns the data model

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sht As Excel.Worksheet
    Dim iRowNum As Integer
    Dim i As Integer
    Dim sSQL As String

    i = 0

    ' SQL like query to get result of DMV from schema $SYSTEM
    sSQL = "select Distinct  [TABLE], [OBJECT], TRIM( '=' +  [EXPRESSION] ) as [DAX Expression]  " & _
             "from $SYSTEM.DISCOVER_CALC_DEPENDENCY  WHERE OBJECT_TYPE = 'CALC_COLUMN'  ORDER BY [TABLE] +[OBJECT]"

    ' Open connection to PowerPivot engine
    Set conn = ActiveWorkbook.Model.DataModelConnection.ModelConnection.ADOConnection
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.Open sSQL, conn, adOpenForwardOnly, adLockOptimistic
        
    If rs.RecordCount > 0 Then
        ReDim aCalcColumns(0 To rs.RecordCount - 1)
        ' Output of the query results
        Do Until rs.EOF
            aCalcColumns(i).Name = rs.Fields("OBJECT")
            aCalcColumns(i).TableName = rs.Fields("TABLE")
            aCalcColumns(i).Expression = rs.Fields("DAX Expression")
            rs.MoveNext
            i = i + 1
        Loop
    End If
    
    rs.Close
    Set rs = Nothing




End Sub


Sub GetModelRelationships(aRelationships() As TypeModelRelationship)

    Dim mdlRelationship As ModelRelationship
    Dim i As Integer
    
    
    i = 0
    ReDim aRelationships(0 To ActiveWorkbook.Model.ModelRelationships.Count - 1)
    For Each mdlRelationship In ActiveWorkbook.Model.ModelRelationships
        aRelationships(i).PrimaryKeyTable = mdlRelationship.PrimaryKeyTable.Name
        aRelationships(i).PrimaryKeyColumn = mdlRelationship.PrimaryKeyColumn.Name
        aRelationships(i).ForeignKeyTable = mdlRelationship.ForeignKeyTable.Name
        aRelationships(i).ForeignKeyColumn = mdlRelationship.ForeignKeyColumn.Name
        aRelationships(i).Active = mdlRelationship.Active
        i = i + 1
    Next mdlRelationship

End Sub

