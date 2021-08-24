Attribute VB_Name = "Reporting_Data"
Attribute VB_Description = "Acts as a point of contact between working modules and underlying storage structure"
'@ModuleDescription "Acts as a point of contact between working modules and underlying storage structure"
'@Folder "Storage.Reporting"
Option Explicit
Option Private Module

Const csReportMetaDataStorageName As String = "ReportProperties"
Const csPivotReportQueriesPerReport As String = "PivotReportingQueriesPerReport"



'---------------------------------------------------------------------------------------
'                               Assign Storage
'---------------------------------------------------------------------------------------

Public Function AssignReportStructureStorage(ByRef wkb As Workbook, _
    Optional bCreateIfNoneExists As Boolean = True) As zLIB_ListStorage

    Dim ls As zLIB_ListStorage
    Dim bStorageIsAssigned As Boolean
    Dim sHeaders(4) As String
    
    Set ls = New zLIB_ListStorage
    bStorageIsAssigned = ls.AssignStorage(wkb, csReportMetaDataStorageName)

    Select Case True
        Case bStorageIsAssigned
            Set AssignReportStructureStorage = ls
        Case Not bStorageIsAssigned And bCreateIfNoneExists
            sHeaders(0) = "ReportType"
            sHeaders(1) = "ReportName"
            sHeaders(2) = "DataType"
            sHeaders(3) = "Property"
            sHeaders(4) = "Value"
            ls.CreateStorage wkb, csReportMetaDataStorageName, sHeaders
            Set AssignReportStructureStorage = ls
        Case Else
            Set AssignReportStructureStorage = Nothing
    End Select

End Function


Public Function AssignPivotTableQueriesPerReport(ByRef wkb As Workbook, _
    Optional bCreateIfNoneExists As Boolean = True) As zLIB_ListStorage

    Dim ls As zLIB_ListStorage
    Dim bStorageIsAssigned As Boolean
    Dim sHeaders(1) As String

    Set ls = New zLIB_ListStorage
    bStorageIsAssigned = ls.AssignStorage(wkb, csPivotReportQueriesPerReport)

    Select Case True
        Case bStorageIsAssigned
            Set AssignPivotTableQueriesPerReport = ls
        Case Not bStorageIsAssigned And bCreateIfNoneExists
            sHeaders(0) = "ReportName"
            sHeaders(1) = "Query"
            ls.CreateStorage wkb, csPivotReportQueriesPerReport, sHeaders
            Set AssignPivotTableQueriesPerReport = ls
        Case Else
            Set AssignPivotTableQueriesPerReport = Nothing
    End Select

End Function



'---------------------------------------------------------------------------------------
'                               Write Data
'---------------------------------------------------------------------------------------


Public Sub DeleteExistingReportData(ByRef vStorageObject As Variant, _
    ByVal sReportType As String, ByVal sReportName As String)
'If there is any existing data in Listorage with same sReportName as report then this is deleted

    Dim sFilterString As String
    Dim ls As zLIB_ListStorage

    Set ls = vStorageObject

    'Filter with different report type or same report type and different name
    sFilterString = "([ReportType] <> ""<ReportType>"")"
    sFilterString = sFilterString & _
          " + (([ReportType] = ""<ReportType>"") * ([ReportName] <> ""<ReportName>""))"
          
    sFilterString = Replace(sFilterString, "<ReportType>", sReportType)
    sFilterString = Replace(sFilterString, "<ReportName>", sReportName)
    ls.Filter sFilterString
    ls.ReplaceDataWithFilteredData

End Sub


Public Sub WriteReportData(ByRef vStorageObject As Variant, sReportType As String, _
    ByVal sReportName As String, ByVal sDataType As String, ByVal DataDictionary As Dictionary)
    
    Dim DataRow As Dictionary
    Dim key As Variant
    Dim ls As zLIB_ListStorage
    
    Set ls = vStorageObject
    For Each key In DataDictionary.Keys
        Set DataRow = New Dictionary
        DataRow.Add "ReportType", sReportType
        DataRow.Add "ReportName", sReportName
        DataRow.Add "DataType", sDataType
        DataRow.Add "Property", key
        DataRow.Add "Value", DataDictionary(key)
        ls.InsertFromDictionary DataRow
        Set DataRow = Nothing
    Next key
    
End Sub




'---------------------------------------------------------------------------------------
'                               Read Data
'---------------------------------------------------------------------------------------

Public Function ReadUniqueSortedReportCategories(ByRef vStorageObject As Variant, _
    Optional ByVal sReportType As String = "") As Variant
'Returns a variant array of unique report categories

    Dim ls As zLIB_ListStorage
    Dim sFilterStr As String

    Set ls = vStorageObject
    If sReportType = "" Then
        sFilterStr = "([DataType] = ""Sheet"") * ([Property] = ""Category"")"
    Else
        sFilterStr = "([DataType] = ""Sheet"") * ([Property] = ""Category"")" & _
            " * ([ReportType] = """ & sReportType & """)"
    End If
    
    ls.Filter sFilterStr

    ReadUniqueSortedReportCategories = ls.ItemsInField( _
        sFieldName:="Value", _
        bUnique:=True, _
        bSorted:=True, _
        SortOrder:=lsAsc, _
        bFiltered:=True)

End Function


Public Function ReadReportNames(ByRef vStorageObject As Variant, _
    Optional ByVal sReportType As String = "", _
    Optional ByVal sCategory As String = "") As Variant
'Returns a variant array of all Reports, optionally filtered by ReportType

    Dim ls As zLIB_ListStorage
    Dim sFilterStr As String

    Set ls = vStorageObject
    'Create a filter string which includes fill population
    sFilterStr = "([ReportName]=[ReportName])"
    
    If sReportType <> "" Then
        sFilterStr = sFilterStr & " * ([ReportType] = """ & sReportType & """)"
    End If

    If sCategory <> "" Then
        sFilterStr = sFilterStr & "([DataType] = ""Sheet"") * ([Property] = ""Category"") * " & _
            "([Value] = """ & sCategory & """)"
    End If

    ls.Filter sFilterStr
    
    ReadReportNames = ls.ItemsInField( _
        sFieldName:="ReportName", _
        bUnique:=True, _
        bSorted:=True, _
        SortOrder:=lsAsc, _
        bFiltered:=True)

End Function


Public Function ReadReportProperties(ByVal vStorageObject As Variant, _
    ByVal sReportType As String, ByVal sReportName As String, _
    ByVal sDataType As String) As Dictionary

    Dim sFilterStr As String
    Dim ls As zLIB_ListStorage
    Dim i As Long
    Dim NumberOfProperties As Long
    Dim Property As String
    Dim Value As String
    Dim ReturnDictionary As Dictionary
    
    Set ReturnDictionary = New Dictionary
    Set ls = vStorageObject
    sFilterStr = "([ReportType]=""" & sReportType & """) * " & _
        "([ReportName]=""" & sReportName & """) * " & _
        "([DataType] = """ & sDataType & """)"
    'Sort by Property (which contains field position) to ensure fields are added in correct order
    ls.Filter sFilterString:=sFilterStr, bSorted:=True, sSortField:="Property", _
        SortOrder:=lsAsc
    NumberOfProperties = ls.NumberOfRecords(bFiltered:=True)
    For i = 1 To NumberOfProperties
        Property = ls.FieldItemByIndex("Property", i, True)
        Value = ls.FieldItemByIndex("Value", i, True)
        ReturnDictionary.Add Property, Value
    Next i

    Set ReadReportProperties = ReturnDictionary

End Function


Public Function ReadQueriesForReportList(ByRef vStorageObject, _
    ByRef ReportList() As TypeReportRecord) As Variant
'Returns the queries to be retained in order to generate ReportList as a
'one dimensional variant array

    Dim ls As zLIB_ListStorage
    Dim sReportListStr As String
    Dim sFilterStr As String
    Dim i As Long
    
    Set ls = vStorageObject
    
    'Compile string "{"ReportName1", "ReportName2", etc}"
    sReportListStr = "{"
    For i = LBound(ReportList) To UBound(ReportList)
        If i = LBound(ReportList) Then
            sReportListStr = sReportListStr & """" & ReportList(i).ReportName & """"
        Else
            sReportListStr = sReportListStr & ",""" & ReportList(i).ReportName & """"
        End If
    Next i
    sReportListStr = sReportListStr & "}"

    sFilterStr = "NOT(ISERROR(XMATCH([ReportName], <ReportNameList>)))"
    sFilterStr = Replace(sFilterStr, "<ReportNameList>", sReportListStr)

    ls.Filter sFilterStr

    ReadQueriesForReportList = ls.ItemsInField("Query", , True, , , True)
    
End Function

