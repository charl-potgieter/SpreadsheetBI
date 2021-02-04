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




Sub CreateDaxQueryTable(ByVal sDaxQueryStr As String, ByRef rng As Range)

    Dim lo As ListObject
    
    'Set source as first connection in the data model.  Seems like a connection needs to be assigned but
    'is irrelevant which table as query is determined by the DAX string anyway
    Set lo = ActiveSheet.ListObjects.Add( _
        SourceType:=xlSrcModel, _
        Source:=ActiveWorkbook.Connections(ActiveWorkbook.Model.ModelTables(1).SourceWorkbookConnection.Name), _
        Destination:=rng)
    
    lo.TableObject.WorkbookConnection.OLEDBConnection.CommandType = xlCmdDAX
    lo.TableObject.WorkbookConnection.OLEDBConnection.CommandText = sDaxQueryStr
    
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

Sub WriteModelMeasuresToSheet()
'Writes model measures to sheet in activeworkbook

    Dim aMeasures() As TypeModelMeasures
    Dim lo As ListObject
    Dim i As Integer

    ' Delete existing sheet if it exists and create new sheet
    If SheetExists(ActiveWorkbook, "ModelMeasures") Then
        ActiveWorkbook.Sheets("ModelMeasures").Delete
    End If
    CreateModelMeasuresSheet ActiveWorkbook
    
    GetModelMeasures ActiveWorkbook, aMeasures
    Set lo = ActiveWorkbook.Sheets("ModelMeasures").ListObjects("tbl_ModelMeasures")
    lo.DataBodyRange.ClearContents
    lo.DataBodyRange.Offset(1, 0).EntireRow.Delete

    If UBound(aMeasures) = 0 And aMeasures(0).Name = "NULL" Then
        GoTo ExitPoint
    End If

    With lo
        For i = 0 To UBound(aMeasures)
            .ListColumns("Name").DataBodyRange.Cells(i + 1) = aMeasures(i).Name
            .ListColumns("Visible").DataBodyRange.Cells(i + 1) = aMeasures(i).Visible
            .ListColumns("Unique Name").DataBodyRange.Cells(i + 1) = aMeasures(i).UniqueName
            .ListColumns("DAX Expression").DataBodyRange.Cells(i + 1) = "':=" & aMeasures(i).Expression
            .ListColumns("Name and Expression").DataBodyRange.Cells(i + 1) = aMeasures(i).Name & ":=" & aMeasures(i).Expression
        Next i
    End With

ExitPoint:

End Sub


Sub WriteModelMeasuresToPipeDelimtedText(ByRef wkb As Workbook, ByVal sFilePathAndName As String)
'Writes model measures to pipe delimited text file

    Dim aMeasures() As TypeModelMeasures
    Dim i As Integer
    Dim sRowToWrite As String
    Dim iFileNo As Integer
    
    GetModelMeasures ActiveWorkbook, aMeasures


    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, """Name""|""Visible""|""Unique Name""|""Dax Expression""|""Name and Expression""";

    If UBound(aMeasures) = 0 And aMeasures(0).Name = "NULL" Then
        GoTo ExitPoint
    End If
    
    For i = 0 To UBound(aMeasures)
        sRowToWrite = vbCrLf & _
            """" & aMeasures(i).Name & """|""" & _
            aMeasures(i).Visible & """|""" & _
            aMeasures(i).UniqueName & """|""" & _
            ":=" & aMeasures(i).Expression & """|""" & _
            aMeasures(i).Name & ":=" & aMeasures(i).Expression & _
            """"
    
        
        Print #iFileNo, sRowToWrite;
    Next i

ExitPoint:
    Close #iFileNo


End Sub



Sub WriteModelMeasuresToHumanReadableText(ByRef wkb As Workbook, ByVal sFilePathAndName As String)
'Writes model measures to pipe delimited text file

    Dim aMeasures() As TypeModelMeasures
    Dim i As Integer
    Dim sRowToWrite As String
    Dim iFileNo As Integer
    
    GetModelMeasures ActiveWorkbook, aMeasures


    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo

    If UBound(aMeasures) = 0 And aMeasures(0).Name = "NULL" Then
        GoTo ExitPoint
    End If
    
    'Write Header
    sRowToWrite = "/********************************************************************** " & vbCrLf & _
        "       DAX measures" & vbCrLf & _
        "***********************************************************************/" & vbCrLf & vbCrLf & vbCrLf & vbCrLf
    Print #iFileNo, sRowToWrite;
    
    
    For i = 0 To UBound(aMeasures)
        If i = 0 Then
            sRowToWrite = ""
        Else
            sRowToWrite = vbCrLf & vbCrLf & vbCrLf
        End If
        sRowToWrite = sRowToWrite & "//--------------------------------------------------------------------" & vbCrLf & _
            "//     " & aMeasures(i).Name & "   (" & aMeasures(i).Table & ")" & vbCrLf & _
            "//-------------------------------------------------------------------- " & vbCrLf & vbCrLf & _
            aMeasures(i).Name & ":=" & aMeasures(i).Expression
        Print #iFileNo, sRowToWrite;
    Next i

ExitPoint:
    Close #iFileNo


End Sub


Sub WriteModelCalcColsToSheet()
'Writes model calculated columns to sheet in activeworkbook
    
    Dim aCalcColumns() As TypeModelCalcColumns
    Dim lo As ListObject
    Dim i As Integer

    ' Delete existing sheet if it exists and create new sheet
    If SheetExists(ActiveWorkbook, "ModelCalcColumns") Then
        ActiveWorkbook.Sheets("ModelCalcColumns").Delete
    End If
    CreateModelCalculatedColumnsSheet ActiveWorkbook

    GetModelCalculatedColumns ActiveWorkbook, aCalcColumns
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

End Sub


Sub WriteModelCalcColsToPipeDelimitedFile(ByRef wkb As Workbook, ByVal sFilePathAndName As String)
'Writes model calculated columns to sheet in activeworkbook
    
    Dim aCalcColumns() As TypeModelCalcColumns
    Dim iFileNo As Integer
    Dim i As Integer
    Dim sRowToWrite As String
    
    GetModelCalculatedColumns ActiveWorkbook, aCalcColumns
    
    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "Name|Table Name|Expression";

    If UBound(aCalcColumns) = 0 And aCalcColumns(0).Name = "NULL" Then
        GoTo ExitPoint
    End If
    
    For i = 0 To UBound(aCalcColumns)
        sRowToWrite = vbCrLf & _
            aCalcColumns(i).Name & "|" & _
            aCalcColumns(i).TableName & "|" & _
            aCalcColumns(i).Expression
            Print #iFileNo, sRowToWrite;
    Next i

ExitPoint:
    Close #iFileNo



End Sub



Sub WriteModelColsToSheet()
'Write model columns to sheet in activeworkbook
    
    Dim aColumns() As TypeModelColumns
    Dim lo As ListObject
    Dim i As Integer

    ' Delete existing sheet if it exists and create new sheet
    If SheetExists(ActiveWorkbook, "ModelColumns") Then
        ActiveWorkbook.Sheets("ModelColumns").Delete
    End If
    CreateModelColumnsSheet ActiveWorkbook

    GetModelColumns ActiveWorkbook, aColumns
    Set lo = ActiveWorkbook.Sheets("ModelColumns").ListObjects("tbl_ModelColumns")
    lo.DataBodyRange.ClearContents
    lo.DataBodyRange.Offset(1, 0).EntireRow.Delete


    If UBound(aColumns) = 0 And aColumns(0).Name = "NULL" Then
        GoTo ExitPoint
    End If

    With lo
        For i = 0 To UBound(aColumns)
            .ListColumns("Name").DataBodyRange.Cells(i + 1) = aColumns(i).Name
            .ListColumns("Table Name").DataBodyRange.Cells(i + 1) = aColumns(i).TableName
            .ListColumns("Unique Name").DataBodyRange.Cells(i + 1) = aColumns(i).UniqueName
            .ListColumns("Visible").DataBodyRange.Cells(i + 1) = aColumns(i).Visible
            .ListColumns("Is calculated column").DataBodyRange.Cells(i + 1).Formula = "=COUNTIFS(tbl_ModelCalcColumns[Name], [@Name]) = 1"
        Next i
    End With

ExitPoint:

End Sub


Sub WriteModelColsToPipeDelimitedFile(ByRef wkb As Workbook, ByVal sFilePathAndName As String)
'Writes model calculated columns to sheet in activeworkbook
    
    Dim aColumns() As TypeModelColumns
    Dim iFileNo As Integer
    Dim i As Integer
    Dim sRowToWrite As String
    
    GetModelColumns ActiveWorkbook, aColumns
    
    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "Name|Table Name|Unique Name|Visible";

    If UBound(aColumns) = 0 And aColumns(0).Name = "NULL" Then
        GoTo ExitPoint
    End If
    
    For i = 0 To UBound(aColumns)
        sRowToWrite = vbCrLf & _
            aColumns(i).Name & "|" & _
            aColumns(i).TableName & "|" & _
            aColumns(i).UniqueName & "|" & _
            aColumns(i).Visible
            Print #iFileNo, sRowToWrite;
    Next i

ExitPoint:
    Close #iFileNo


End Sub





Sub WriteModelRelationshipsToSheet()
'Write model relationships to sheet in activeworkbook
    
    Dim aModelRelationships() As TypeModelRelationship
    Dim lo As ListObject
    Dim i As Integer

    ' Delete existing sheet if it exists and create new sheet
    If SheetExists(ActiveWorkbook, "ModelRelationships") Then
        ActiveWorkbook.Sheets("ModelRelationships").Delete
    End If
    CreateModelRelationshipsSheet ActiveWorkbook

    GetModelRelationships ActiveWorkbook, aModelRelationships
    Set lo = ActiveWorkbook.Sheets("ModelRelationships").ListObjects("tbl_ModelRelationships")
    lo.DataBodyRange.ClearContents
    lo.DataBodyRange.Offset(1, 0).EntireRow.Delete

    If UBound(aModelRelationships) = 0 And aModelRelationships(0).ForeignKeyColumn = "NULL" Then
        GoTo ExitPoint
    End If

    With lo
        For i = 0 To UBound(aModelRelationships)
            .ListColumns("Primary Key Table").DataBodyRange.Cells(i + 1) = aModelRelationships(i).PrimaryKeyTable
            .ListColumns("Primary Key Column").DataBodyRange.Cells(i + 1) = aModelRelationships(i).PrimaryKeyColumn
            .ListColumns("Foreign Key Table").DataBodyRange.Cells(i + 1) = aModelRelationships(i).ForeignKeyTable
            .ListColumns("Foreign Key Column").DataBodyRange.Cells(i + 1) = aModelRelationships(i).ForeignKeyColumn
            .ListColumns("Active").DataBodyRange.Cells(i + 1) = aModelRelationships(i).Active
        Next i
    End With

ExitPoint:
    

End Sub


Sub WriteModelRelationshipsToPipeDelimitedFile(ByRef wkb As Workbook, ByVal sFilePathAndName As String)
'Writes model calculated columns to sheet in activeworkbook
    
    Dim aRelationships() As TypeModelRelationship
    Dim iFileNo As Integer
    Dim i As Integer
    Dim sRowToWrite As String
    
    GetModelRelationships ActiveWorkbook, aRelationships
    
    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "Primary Key Table|Primary Key Column|Foreign Key Table|Foreign Key Column|Active";

    If UBound(aRelationships) = 0 And aRelationships(0).PrimaryKeyColumn = "NULL" Then
        GoTo ExitPoint
    End If
    
    For i = 0 To UBound(aRelationships)
        sRowToWrite = vbCrLf & _
            aRelationships(i).PrimaryKeyTable & "|" & _
            aRelationships(i).PrimaryKeyColumn & "|" & _
            aRelationships(i).ForeignKeyTable & "|" & _
            aRelationships(i).ForeignKeyColumn & "|" & _
            aRelationships(i).Active
            Print #iFileNo, sRowToWrite;
    Next i

ExitPoint:
    Close #iFileNo


End Sub


Sub test()

    WriteModelRelationshipsToPipeDelimitedFile Workbooks("Temp.xlsx"), "C:\Users\charl\Dropbox\Dropbox_Charl\Computer_Technical\Programming_GitHub\SpreadsheetBI\TestRelationships.txt"

End Sub


Sub GetModelMeasures(ByRef wkb As Workbook, ByRef aModelMeasures() As TypeModelMeasures)
'Requires reference to Microsoft ActiveX Data Objects
'Returns measures in the data model
'If no measures exist a single record is returned with a single record of "NULL" string values

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sht As Excel.Worksheet
    Dim iRowNum As Integer
    Dim i As Integer
    Dim sSQL As String

    i = 0

    ' SQL like query to get result of DMV from schema $SYSTEM
     sSQL = "select [MEASURE_NAME], [MEASURE_UNIQUE_NAME], [MEASURE_IS_VISIBLE], [EXPRESSION], [MEASUREGROUP_NAME] from $SYSTEM.MDSCHEMA_MEASURES  " & _
            "WHERE LEN([EXPRESSION]) > 0 AND " & _
            "[MEASURE_NAME] <> '__No measures defined' " & _
            "ORDER BY [MEASUREGROUP_NAME]"
    

    ' Open connection to PowerPivot engine
    Set conn = wkb.Model.DataModelConnection.ModelConnection.ADOConnection
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
            aModelMeasures(i).Table = rs.Fields("MEASUREGROUP_NAME")
            rs.MoveNext
            i = i + 1
        Loop
    Else
        ReDim aModelMeasures(0 To 0)
        aModelMeasures(0).Name = "NULL"
        aModelMeasures(0).Expression = "NULL"
        aModelMeasures(0).UniqueName = "NULL"
        aModelMeasures(0).Visible = False
    End If
    
    rs.Close
    Set rs = Nothing


End Sub


Sub GetModelColumns(ByRef wkb As Workbook, ByRef aModelColumns() As TypeModelColumns)
'Requires reference to Microsoft ActiveX Data Objects
'Returns columns the data model
'If no columns exist a single record is returned with a single record of "NULL" string values

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
    Set conn = wkb.Model.DataModelConnection.ModelConnection.ADOConnection
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
    Else
        ReDim aModelColumns(0 To 0)
        aModelColumns(0).Name = "NULL"
        aModelColumns(0).TableName = "NULL"
        aModelColumns(0).UniqueName = "NULL"
        aModelColumns(0).Visible = False
    End If
    
    rs.Close
    Set rs = Nothing


End Sub

Sub GetModelCalculatedColumns(ByRef wkb As Workbook, ByRef aCalcColumns() As TypeModelCalcColumns)
'Requires reference to Microsoft ActiveX Data Objects
'Returns calcualted columns the data model
'If no calculated columns exist a single record is returned with a single record of "NULL" string values

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
    Set conn = wkb.Model.DataModelConnection.ModelConnection.ADOConnection
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
    Else
        ReDim aCalcColumns(0 To 0)
        aCalcColumns(0).Name = "NULL"
        aCalcColumns(0).TableName = "NULL"
        aCalcColumns(0).Expression = "NULL"
    End If
    
    rs.Close
    Set rs = Nothing




End Sub


Sub GetModelRelationships(ByRef wkb As Workbook, aRelationships() As TypeModelRelationship)
'Requires reference to Microsoft ActiveX Data Objects
'Returns data model relationships
'If no relationships exist a single record is returned with a single record of "NULL" string values

    Dim mdlRelationship As ModelRelationship
    Dim i As Integer
    
    
    i = 0
    If wkb.Model.ModelRelationships.Count > 0 Then
        ReDim aRelationships(0 To wkb.Model.ModelRelationships.Count - 1)
        For Each mdlRelationship In wkb.Model.ModelRelationships
            aRelationships(i).PrimaryKeyTable = mdlRelationship.PrimaryKeyTable.Name
            aRelationships(i).PrimaryKeyColumn = mdlRelationship.PrimaryKeyColumn.Name
            aRelationships(i).ForeignKeyTable = mdlRelationship.ForeignKeyTable.Name
            aRelationships(i).ForeignKeyColumn = mdlRelationship.ForeignKeyColumn.Name
            aRelationships(i).Active = mdlRelationship.Active
            i = i + 1
        Next mdlRelationship
    Else
        ReDim aRelationships(0 To 0)
        aRelationships(0).PrimaryKeyTable = "NULL"
        aRelationships(0).PrimaryKeyColumn = "NULL"
        aRelationships(0).ForeignKeyTable = "NULL"
        aRelationships(0).ForeignKeyColumn = "NULL"
        aRelationships(0).Active = False
    End If

End Sub





Sub CreateModelMeasuresSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject

    
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ModelMeasures"
    sht.Range("SheetHeading") = "Data model measures"
    sht.Range("SheetCategory") = "Setup"
   
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:F7"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_ModelMeasures"
        .HeaderRowRange.Cells(1) = "Name"
        .HeaderRowRange.Cells(2) = "Visible"
        .HeaderRowRange.Cells(3) = "Unique Name"
        .HeaderRowRange.Cells(4) = "DAX Expression"
        .HeaderRowRange.Cells(5) = "Name and Expression"
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 40
    sht.Range("C:C").ColumnWidth = 20
    sht.Range("D:D").ColumnWidth = 40
    sht.Range("E:E").ColumnWidth = 80
    sht.Range("F:F").ColumnWidth = 80

    With lo.DataBodyRange
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With

    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 6
    ActiveWindow.FreezePanes = True

End Sub


Sub CreateModelColumnsSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject

   
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ModelColumns"
    sht.Range("SheetHeading") = "Data model columns"
    sht.Range("SheetCategory") = "Setup"
    sht.Range("B4") = "Includes calculated columns"
   
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:F7"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_ModelColumns"
        .HeaderRowRange.Cells(1) = "Name"
        .HeaderRowRange.Cells(2) = "Table Name"
        .HeaderRowRange.Cells(3) = "Unique Name"
        .HeaderRowRange.Cells(4) = "Visible"
        .HeaderRowRange.Cells(5) = "Is calculated column"
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 30
    sht.Range("C:C").ColumnWidth = 30
    sht.Range("D:D").ColumnWidth = 50
    sht.Range("E:E").ColumnWidth = 20
    sht.Range("F:F").ColumnWidth = 20

    With lo.DataBodyRange
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With

    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 6
    ActiveWindow.FreezePanes = True


End Sub






Sub CreateModelCalculatedColumnsSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject

    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ModelCalcColumns"
    sht.Range("SheetHeading") = "Data model calculated columns"
    sht.Range("SheetCategory") = "Setup"
   
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:D7"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_ModelCalcColumns"
        .HeaderRowRange.Cells(1) = "Name"
        .HeaderRowRange.Cells(2) = "Table Name"
        .HeaderRowRange.Cells(3) = "Expression"
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 30
    sht.Range("C:C").ColumnWidth = 30
    sht.Range("D:D").ColumnWidth = 50

    With lo.DataBodyRange
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With

    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 6
    ActiveWindow.FreezePanes = True


End Sub



Sub CreateModelRelationshipsSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject

    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "ModelRelationships"
    sht.Range("SheetHeading") = "Data model relationships"
    sht.Range("SheetCategory") = "Setup"
   
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B6:F7"), XlListObjectHasHeaders:=xlYes)
    With lo
        .Name = "tbl_ModelRelationships"
        .HeaderRowRange.Cells(1) = "Primary Key Table"
        .HeaderRowRange.Cells(2) = "Primary Key Column"
        .HeaderRowRange.Cells(3) = "Foreign Key Table"
        .HeaderRowRange.Cells(4) = "Foreign Key Column"
        .HeaderRowRange.Cells(5) = "Active"
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 40
    sht.Range("C:C").ColumnWidth = 40
    sht.Range("D:D").ColumnWidth = 40
    sht.Range("E:E").ColumnWidth = 40
    sht.Range("F:F").ColumnWidth = 20

    With lo.DataBodyRange
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With

    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 6
    ActiveWindow.FreezePanes = True


End Sub




Sub CreateTableGeneratorSheet(ByRef wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject

    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    FormatSheet sht
    sht.Name = "PqTableGenerator"
    sht.Range("SheetHeading") = "Power query table generator"
    sht.Range("SheetCategory") = "Setup"
   
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("B11:F12"), XlListObjectHasHeaders:=xlYes)
    FormatTable lo
    
    With lo
        .Name = "tbl_PqTableGenerator"
        .HeaderRowRange.Cells(1) = "Column_1"
        .HeaderRowRange.Cells(2) = "Column_2"
        .HeaderRowRange.Cells(3) = "Column_3"
        .HeaderRowRange.Cells(4) = "Column_4"
        .HeaderRowRange.Cells(5) = "Column_5"
    End With
    FormatTable lo
    sht.Range("B:B").ColumnWidth = 20
    sht.Range("C:C").ColumnWidth = 20
    sht.Range("D:D").ColumnWidth = 20
    sht.Range("E:E").ColumnWidth = 20
    sht.Range("F:F").ColumnWidth = 20


    'Add various formatted text to the sheet
    sht.Range("B5") = "Generates a power query with hardcoded values and field types as below, using the GeneratePowerQuery code"
    sht.Range("B7") = "Query Name"
    sht.Range("C7") = "TestTable"
    sht.Range("B7").Font.Bold = True
    sht.Range("C7,B9:F9").Interior.Color = RGB(242, 242, 242)
    sht.Range("C7,B9:F9").Font.Color = RGB(0, 112, 192)
    sht.Range("B9:F9").HorizontalAlignment = xlCenter
    sht.Range("B9:F9") = "type text"
    
    'Add data validation for field types
    sht.Range("B9:F9").Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:="type any,type binary,type date,type datetime,type datetimezone,type duration,Int64.Type," & _
            "type logical,type none,type number,type text,type time"
  
    'Create Named Range
    sht.Names.Add Name:="TableName", RefersToR1C1:="=R7C3"
  

    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 11
    ActiveWindow.FreezePanes = True


End Sub



