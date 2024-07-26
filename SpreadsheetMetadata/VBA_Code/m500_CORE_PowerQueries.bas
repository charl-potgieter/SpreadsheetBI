Attribute VB_Name = "m500_CORE_PowerQueries"
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


Sub ExportPowerQueriesToConsolidatedFile(ByVal wkb As Workbook)

    Dim qry As WorkbookQuery
    Dim ExportFileNameAndPath As String
    Dim QueryString As String
    Dim isFirstQuery As Boolean
    
    ExportFileNameAndPath = wkb.Path & Application.PathSeparator & "ConsolidatedPowerQueries.m"
    isFirstQuery = True
    
    QueryString = "[" & vbLf
    For Each qry In wkb.Queries
        If isFirstQuery Then
            QueryString = QueryString & vbLf & vbLf & qry.Name & " = " & vbLf
        Else
            QueryString = QueryString & "," & vbLf & vbLf & qry.Name & " = " & vbLf
        End If
        QueryString = QueryString & qry.Formula
        isFirstQuery = False
    Next qry
    QueryString = QueryString & "]"

    WriteStringToTextFile QueryString, ExportFileNameAndPath

End Sub


Sub GenerateConsolidatedPowerQuery(ByVal wkb As Workbook, ByVal ConsolQryName As String)

    Dim qry As WorkbookQuery
    Dim ExportFileNameAndPath As String
    Dim QueryString As String
    Dim isFirstQuery As Boolean
    
    ExportFileNameAndPath = wkb.Path & Application.PathSeparator & "ConsolidatedPowerQueries.m"
    isFirstQuery = True
    
    QueryString = "[" & vbLf
    For Each qry In wkb.Queries
        If qry.Name <> ConsolQryName And Left(qry.Name, 1) <> "_" Then
            If isFirstQuery Then
                QueryString = QueryString & vbLf & vbLf & qry.Name & " = " & vbLf
            Else
                QueryString = QueryString & "," & vbLf & vbLf & qry.Name & " = " & vbLf
            End If
            QueryString = QueryString & qry.Formula
            isFirstQuery = False
        End If
    Next qry
    QueryString = QueryString & "]"
    
    If QueryExists(ConsolQryName) Then
        wkb.Queries(ConsolQryName).Formula = QueryString
    Else
        wkb.Queries.Add ConsolQryName, QueryString
    End If


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



Sub GetPowerQueryFileNamesFromUser(ByRef FilePaths() As String)

    Dim sPowerQueryFilePath As String
    Dim sPowerQueryName As String
    Dim fDialog As FileDialog
    Dim fso As FileSystemObject
    Dim i As Integer
    Dim LatestSelectedFilePath As String

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

    With fDialog
        .AllowMultiSelect = True
        .Title = "Select power query / queries"
        .InitialFileName = ThisWorkbook.Path
        .Filters.Clear
        .Filters.Add "m Power Query Files", "*.m"
        .InitialFileName = GetSundryStorageItem("Last used directory text referenced power queries")
    End With

    'fDialog.Show value of -1 below means success
    If fDialog.Show = -1 Then
        ReDim Preserve FilePaths(0 To fDialog.SelectedItems.Count - 1)
        For i = 0 To fDialog.SelectedItems.Count - 1
            FilePaths(i) = fDialog.SelectedItems(i + 1)
        Next i
    End If

    Set fso = New FileSystemObject
    LatestSelectedFilePath = fso.GetParentFolderName(fDialog.SelectedItems(1))
    UpdateSundryStorageValueForGivenItem "Last used directory text referenced power queries", _
        LatestSelectedFilePath
    
    
    ThisWorkbook.Save

End Sub


Function PowerQueryReferencedToTextFile(ByVal FileName As String) As String

    PowerQueryReferencedToTextFile = _
        "let" & vbLf & _
        "   FileName = ""<FileName>""," & vbLf & _
        "   Binary = File.Contents(FileName)," & vbLf & _
        "   QueryText = Text.FromBinary(Binary)," & vbLf & _
        "   Output = Expression.Evaluate(QueryText, #shared)" & vbLf & _
        "in" & vbLf & _
        "   Output"

    PowerQueryReferencedToTextFile = Replace(PowerQueryReferencedToTextFile, _
        "<FileName>", FileName)

End Function


Function GetConnectionFromPowerQueryName(ByVal wkb As Workbook, ByVal PowerQueryName As String) As WorkbookConnection

    Dim cn As WorkbookConnection
    Dim i As Integer
    Dim MatchFound As Boolean
    
    MatchFound = False
    i = 1
    
    Do While i <= wkb.Connections.Count And Not MatchFound
        Set cn = wkb.Connections(i)
        If cn.Type = xlConnectionTypeOLEDB Then
            If cn.OLEDBConnection.CommandText = """" & PowerQueryName & """" Then
                Set GetConnectionFromPowerQueryName = cn
                MatchFound = True
            End If
        End If
        i = i + 1
    Loop
    
End Function


Sub LoadPowerQueryToTable(ByVal TargetSht As Worksheet, ByVal QueryName As String)

    Dim QryTable As QueryTable
    Dim lo As ListObject
    Dim SourceString As String
    Dim CommandString As String

    SourceString = "OLEDB;" & _
        "Provider=Microsoft.Mashup.OleDb.1;" & _
        "Data Source=$Workbook$;" & _
        "Location=" & QueryName & ";" & _
        "Extended Properties="""""
        
    CommandString = "SELECT * FROM [" & QueryName & "]"
    
    Set lo = TargetSht.ListObjects.Add( _
        SourceType:=xlSrcExternal, _
        Source:=SourceString, _
        Destination:=Range("$A$1"))
    
    
    
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array(CommandString)
        .Refresh BackgroundQuery:=False
    End With
    
End Sub



Function QueryExists(ByVal sQryName As String, Optional wkb As Workbook) As Boolean

    If wkb Is Nothing Then
        Set wkb = ActiveWorkbook
    End If
    
    On Error Resume Next
    QueryExists = CBool(Len(wkb.Queries(sQryName).Name))
    On Error GoTo 0
    
End Function
