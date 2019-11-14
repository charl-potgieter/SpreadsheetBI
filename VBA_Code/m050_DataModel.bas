Attribute VB_Name = "m050_DataModel"
Option Explicit
Option Private Module

Sub ExportPowerQueriesToFiles(ByVal sFolderPath As String, wkb As Workbook)

    Dim qry As WorkbookQuery
    
    For Each qry In wkb.Queries
        WriteStringToTextFile qry.Formula, sFolderPath & Application.PathSeparator & qry.Name & ".m"
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
    
    wkb.Model.Add _
        MeasureName:=sMeasureName, _
        AssociatedTable:=wkb.Model.ModelTables(sTableName), _
        Formula:=sMeasure, _
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

