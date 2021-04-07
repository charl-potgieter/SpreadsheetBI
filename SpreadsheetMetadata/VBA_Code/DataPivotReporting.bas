Attribute VB_Name = "DataPivotReporting"
Attribute VB_Description = "Acts as a point of contact between working modules and underlying storage structure"
'@ModuleDescription "Acts as a point of contact between working modules and underlying storage structure"
'@Folder "Reporting"
Option Explicit
Option Private Module

Const csPivotReportMetaDataStorageName As String = "PivotReportProperties"
Const csPivotReportExcelTableFormulas As String = "PivotReportExcelTableFormulas"



'---------------------------------------------------------------------------------------
'                               Assign Storage
'---------------------------------------------------------------------------------------

Public Function AssignPivotReportStructureStorage(ByRef wkb As Workbook, _
    Optional bCreateIfNoneExists As Boolean = False) As ListStorage
'Assigns storage to liststorage object if it exists.  Storage is created if none exits
'and bCreateIfNoneExists is set as True

    Dim ls As ListStorage
    Dim bStorageIsAssigned As Boolean
    Dim sHeaders(5) As String
    
    Set ls = New ListStorage
    bStorageIsAssigned = ls.AssignStorage(wkb, csPivotReportMetaDataStorageName)

    Select Case True
        Case bStorageIsAssigned
            Set AssignPivotReportStructureStorage = ls
        Case Not bStorageIsAssigned And bCreateIfNoneExists
            sHeaders(0) = "ReportName"
            sHeaders(1) = "Name"
            sHeaders(2) = "DataType"
            sHeaders(3) = "Property"
            sHeaders(4) = "Value"
            ls.CreateStorage wkb, csPivotReportMetaDataStorageName, sHeaders
            Set AssignPivotReportStructureStorage = ls
        Case Else
            Set AssignPivotReportStructureStorage = Nothing
    End Select

End Function


Public Function AssignPivotReportFormulaStorage(ByRef wkb As Workbook) As ListStorage
'Create a ListStorage sheet for storing queries that can optionaly be used to provide data backing
'PowerReports (per PowerReport class module)

    Dim ls As ListStorage
    Dim bStorageIsAssigned As Boolean
    Dim sHeaders(1) As String

    Set ls = New ListStorage
    bStorageIsAssigned = ls.AssignStorage(wkb, csPivotReportExcelTableFormulas)

    If Not bStorageIsAssigned Then
        sHeaders(0) = "ReportName"
        sHeaders(1) = "Formula"
        ls.CreateStorage wkb, csPivotReportExcelTableFormulas, sHeaders
    End If

    Set AssignPivotReportFormulaStorage = ls

End Function



'---------------------------------------------------------------------------------------
'                               Write Data
'---------------------------------------------------------------------------------------






Public Sub DeleteExistingPivotReportData(ByRef vStorageObject As Variant, _
    ByVal sReportName As String)
'If there is any existing data in Listorage with same sReportName as report then this is deleted

    Dim sFilterString As String
    Dim ls As ListStorage

    Set ls = vStorageObject

    'Create Filter excluding Report name and then replace data with filter
    sFilterString = "[ReportName] <> """ & sReportName & """"
    ls.Filter sFilterString
    ls.ReplaceDataWithFilteredData

End Sub


Public Sub WritePivotReportData(ByRef vStorageObject As Variant, ByVal sReportName As String, _
    ByVal sDataType As String, ByVal DataDictionary As Dictionary)
    
    Dim DataRow As Dictionary
    Dim key As Variant
    Dim ls As ListStorage
    
    Set ls = vStorageObject
    For Each key In DataDictionary.Keys
        Set DataRow = New Dictionary
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

Public Function ReadUniquePivotReportCategories(ByRef vStorageObject As Variant) As Variant
'Returns a variant array of unique report categories

    Dim ls As ListStorage

    Set ls = vStorageObject
    ls.Filter "([DataType] = ""Sheet"") * ([Property] = ""Category"")"

    ReadUniquePivotReportCategories = ls.ItemsInField( _
        sFieldName:="Value", _
        bUnique:=True, _
        bSorted:=True, _
        SortOrder:=lsAsc, _
        bFiltered:=True)

End Function



Public Function ReadAllPivotReports(ByRef vStorageObject As Variant) As Variant
'Returns a variant array of all PowerrReports

    Dim ls As ListStorage

    Set ls = vStorageObject

    ReadAllPivotReports = ls.ItemsInField( _
        sFieldName:="ReportName", _
        bUnique:=True, _
        bSorted:=True, _
        SortOrder:=lsAsc, _
        bFiltered:=False)

End Function



Public Function ReadPivotReportsByCategory(ByRef vStorageObject, sReportCategory As String) As Variant
'Returns a variant array of PowerReports based sReportCategory

    Dim ls As ListStorage
    Dim sFilterStr As String

    Set ls = vStorageObject

    sFilterStr = "([DataType] = ""Sheet"") * ([Property] = ""Category"") * " & _
        "([Value] = """ & sReportCategory & """)"
    ls.Filter sFilterStr

    ReadPivotReportsByCategory = ls.ItemsInField( _
        sFieldName:="ReportName", _
        bUnique:=True, _
        bSorted:=True, _
        SortOrder:=lsAsc, _
        bFiltered:=True)

End Function



Public Function ReadPivotReportHasExcelTableSource(ByRef vStorageObject As Variant, _
    ByVal sReportName As String) As Boolean

    Dim ls As ListStorage
    Set ls = vStorageObject

    ReadPivotReportHasExcelTableSource = CBool(ls.Xlookup(sReportName & "Sheet" & "HasExcelTableSource", _
    "[ReportName] & [DataType] & [Property]", "[Value]"))

End Function


Public Function ReadPivotReportProperties(ByVal vStorageObject As Variant, _
    ByVal PvtReport As PivotReport, ByVal sDataType As String) As Dictionary

    Dim sFilterStr As String
    Dim ls As ListStorage
    Dim i As Long
    Dim NumberOfProperties As Long
    Dim Property As String
    Dim Value As String
    Dim ReturnDictionary As Dictionary
    
    Set ReturnDictionary = New Dictionary
    Set ls = vStorageObject
    sFilterStr = "([ReportName]=""" & PvtReport.ReportName & """) * " & _
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

    Set ReadPivotReportProperties = ReturnDictionary

End Function


'Public Function ReadCubeFieldNames(ByVal vStorageObject As Variant, _
'    ByVal PvtReport As PivotReport) As Variant
''Returns a one-dimensional base 0 array of CubeFieldNames
'
'    Dim sFilterStr As String
'    Dim ls As ListStorage
'
'
'    Set ls = vStorageObject
'    sFilterStr = "([ReportName] = ""<<Report Name>>"") * ([DataType] = ""CubeFieldOrientation"")"
'    sFilterStr = Replace(sFilterStr, "<<Report Name>>", PvtReport.ReportName)
'    ls.Filter sFilterStr
'    Set ReadCubeFieldNames = ls.ItemsInField(sFieldName:="Name", bUnique:=True, bFiltered:=True)
'
'End Function
'



'Sub PR_WriteRecordToStorage( _
'    ByRef vStorageObject As Variant, _
'    ByRef StorageRecords() As TypePowerReportStorageRecord)
''Writes a single property of PowerReport to list storage
''Deletes any previous data that exists for a report
'
'    Dim dict As Dictionary
'    Dim ls As ListStorage
'    Dim i As Long
'
'    Set ls = vStorageObject
'
'    For i = LBound(StorageRecords) To UBound(StorageRecords)
'        Set dict = New Dictionary
'        dict.Add "ReportName", StorageRecords(i).ReportName
'        dict.Add "Name", StorageRecords(i).Name
'        dict.Add "DataType", StorageRecords(i).DataType
'        dict.Add "Property", StorageRecords(i).Property
'        dict.Add "Value", StorageRecords(i).Value
'        dict.Add "CubeFieldPosition", StorageRecords(i).CubeFieldPosition
'        ls.InsertFromDictionary dict
'    Next i
'
'End Sub
'
'Function PR_GetHeadingBasedOnReportName( _
'    ByRef vStorageObject As Variant, ByVal sReportName As String) As String
'
'    Dim ls As ListStorage
'
'    Set ls = vStorageObject
'
'    PR_GetHeadingBasedOnReportName = ls.Xlookup( _
'        LookupValue:=sReportName & "SheetDataType" & "SheetHeading", _
'        sLookupArray:="[ReportName] & [DataType] & [Property]", _
'        sReturnArray:="[Value]")
'
'End Function
'
'
'Function PR_GetCategoryBasedOnReportName( _
'    ByRef vStorageObject As Variant, ByVal sReportName As String) As String
'
'    Dim ls As ListStorage
'
'    Set ls = vStorageObject
'
'    PR_GetCategoryBasedOnReportName = ls.Xlookup( _
'        LookupValue:=sReportName & "Sheet" & "SheetCategory", _
'        sLookupArray:="[ReportName] & [DataType] & [Property]", _
'        sReturnArray:="[Value]")
'
'End Function
'
'Function PR_GetPivotTableProperties(ByRef vStorageObject As Variant, _
'    ByVal sReportName As String, _
'    ByRef StorageRecords() As TypePowerReportStorageRecord)
''Returns  Pivot Table properties in the Storage Array.
'
'
'    Dim ls As ListStorage
'    Dim i As Long
'    Dim sFilterStr As String
'
'    Set ls = vStorageObject
'
'    sFilterStr = "([ReportName]=""" & sReportName & """) * " & _
'        "([DataType] = ""PivotTable"")"
'
'    ls.Filter sFilterStr
'
'    ReDim StorageRecords(ls.NumberOfRecords(bFiltered:=True) - 1)
'
'    For i = LBound(StorageRecords) To UBound(StorageRecords)
'        StorageRecords(i).ReportName = ls.FieldItemByIndex("ReportName", i + 1, True)
'        StorageRecords(i).Name = ls.FieldItemByIndex("Name", i + 1, True)
'        StorageRecords(i).DataType = ls.FieldItemByIndex("DataType", i + 1, True)
'        StorageRecords(i).Property = ls.FieldItemByIndex("Property", i + 1, True)
'        StorageRecords(i).Value = ls.FieldItemByIndex("Value", i + 1, True)
'    Next i
'
'End Function
'
'
'Function PR_GetPivotCubeFieldDataOrientationSortedByCubeFieldPosition( _
'    ByRef vStorageObject As Variant, ByVal sReportName As String, _
'    ByRef StorageRecords() As TypePowerReportStorageRecord)
''Returns storage orientation property of pivot cube fields in the
''Storage Array.   Sorted by CubefieldPosition to ensure correct position when pivot table is created
''For some reason it doesn't work setting the position property directly in VBA
'
'
'    Dim ls As ListStorage
'    Dim i As Long
'    Dim sFilterStr As String
'
'    Set ls = vStorageObject
'
'    sFilterStr = "([ReportName]=""" & sReportName & """) * " & _
'        "([DataType] = ""PivotCubeField"") * " & _
'        "([Property] = ""Orientation"")"
'
'    ls.Filter sFilterStr, True, "[CubeFieldPosition]", lsAsc
'
'    ReDim StorageRecords(ls.NumberOfRecords(bFiltered:=True) - 1)
'
'    For i = LBound(StorageRecords) To UBound(StorageRecords)
'        StorageRecords(i).ReportName = ls.FieldItemByIndex("ReportName", i + 1, True)
'        StorageRecords(i).Name = ls.FieldItemByIndex("Name", i + 1, True)
'        StorageRecords(i).DataType = ls.FieldItemByIndex("DataType", i + 1, True)
'        StorageRecords(i).Property = ls.FieldItemByIndex("Property", i + 1, True)
'        StorageRecords(i).Value = ls.FieldItemByIndex("Value", i + 1, True)
'    Next i
'
'End Function
'
'
'Function PR_GetPivotCubeFieldDataPropertiesExOrientation( _
'    ByRef vStorageObject As Variant, ByVal sReportName As String, _
'    ByRef StorageRecords() As TypePowerReportStorageRecord)
''Returns storage properties ex-orientation property of pivot cube fields to the
''Storage Array
'
'
'    Dim ls As ListStorage
'    Dim i As Long
'
'    Set ls = vStorageObject
'
'    ls.Filter "([ReportName]=""" & sReportName & """) * " & _
'        "([DataType] = ""PivotCubeField"") * " & _
'        "([Property] <> ""Orientation"")"
'
'    ReDim StorageRecords(ls.NumberOfRecords(bFiltered:=True) - 1)
'
'    For i = LBound(StorageRecords) To UBound(StorageRecords)
'        StorageRecords(i).ReportName = ls.FieldItemByIndex("ReportName", i + 1, True)
'        StorageRecords(i).Name = ls.FieldItemByIndex("Name", i + 1, True)
'        StorageRecords(i).DataType = ls.FieldItemByIndex("DataType", i + 1, True)
'        StorageRecords(i).Property = ls.FieldItemByIndex("Property", i + 1, True)
'        StorageRecords(i).Value = ls.FieldItemByIndex("Value", i + 1, True)
'    Next i
'
'End Function
'
'Function PR_GetPivotFieldDataSubtotalProperty( _
'    ByRef vStorageObject As Variant, ByVal sReportName As String, _
'    ByRef StorageRecords() As TypePowerReportStorageRecord)
''Returns storage properties of pivot field subtotal to the Storage Array
''Subtotal handled seperately as it has flow on effects and slightly different in that property
''needs to be indexed
'
'
'    Dim ls As ListStorage
'    Dim i As Long
'
'    Set ls = vStorageObject
'
'    ls.Filter "([ReportName]=""" & sReportName & """) * " & _
'        "([DataType] = ""PivotField"") * " & _
'        "([Property] = ""Subtotals"")"
'
'    ReDim StorageRecords(ls.NumberOfRecords(bFiltered:=True) - 1)
'
'    For i = LBound(StorageRecords) To UBound(StorageRecords)
'        StorageRecords(i).ReportName = ls.FieldItemByIndex("ReportName", i + 1, True)
'        StorageRecords(i).Name = ls.FieldItemByIndex("Name", i + 1, True)
'        StorageRecords(i).DataType = ls.FieldItemByIndex("DataType", i + 1, True)
'        StorageRecords(i).Property = ls.FieldItemByIndex("Property", i + 1, True)
'        StorageRecords(i).Value = ls.FieldItemByIndex("Value", i + 1, True)
'    Next i
'
'End Function
'
'
'Function PR_GetPivotFieldDataPropertiesExSubtotals( _
'    ByRef vStorageObject As Variant, ByVal sReportName As String, _
'    ByRef StorageRecords() As TypePowerReportStorageRecord)
''Returns storage properties of pivot fields to the Storage Array
'
'
'    Dim ls As ListStorage
'    Dim i As Long
'
'    Set ls = vStorageObject
'
'    ls.Filter "([ReportName]=""" & sReportName & """) * " & _
'        "([DataType] = ""PivotField"") * " & _
'        "([Property] <> ""Subtotals"")"
'
'    ReDim StorageRecords(ls.NumberOfRecords(bFiltered:=True) - 1)
'
'    For i = LBound(StorageRecords) To UBound(StorageRecords)
'        StorageRecords(i).ReportName = ls.FieldItemByIndex("ReportName", i + 1, True)
'        StorageRecords(i).Name = ls.FieldItemByIndex("Name", i + 1, True)
'        StorageRecords(i).DataType = ls.FieldItemByIndex("DataType", i + 1, True)
'        StorageRecords(i).Property = ls.FieldItemByIndex("Property", i + 1, True)
'        StorageRecords(i).Value = ls.FieldItemByIndex("Value", i + 1, True)
'    Next i
'
'End Function
'
'
'Function PR_GetFreezePaneLocation(ByRef vStorageObject As Variant, ByVal sReportName As String) As Variant
''Returns null if none found
'
'    Dim ls As ListStorage
'    Dim sFreezePosition As String
'
'    Set ls = vStorageObject
'
'    PR_GetFreezePaneLocation = ls.Xlookup(sReportName & "ViewLayout" & "FreezePanes", _
'        "[ReportName] & [DataType] & [Property]", _
'        "[Value]")
'
'End Function
'
'
'Function PR_GetRowRangeColWidths(ByRef vStorageObject As Variant, _
'    ByVal sReportName As String) As Variant
''Returns null if none found, else pipe delimited string of widths
'
'    Dim ls As ListStorage
'    Dim sFreezePosition As String
'
'    Set ls = vStorageObject
'
'    PR_GetRowRangeColWidths = ls.Xlookup(sReportName & "ViewLayout" & "PivotRowRangeColWidths", _
'        "[ReportName] & [DataType] & [Property]", _
'        "[Value]")
'
'End Function
'
'
'Function PR_GetDataBodyRowRangeColWidth(ByRef vStorageObject As Variant, _
'    ByVal sReportName As String) As Variant
''Returns null if none found, else pipe delimited string of widths
'
'    Dim ls As ListStorage
'    Dim sFreezePosition As String
'
'    Set ls = vStorageObject
'
'    PR_GetDataBodyRowRangeColWidth = ls.Xlookup(sReportName & "ViewLayoute" & "PivotDataBodyRangeColWidths", _
'        "[ReportName] & [DataType] & [Property]", _
'        "[Value]")
'
'End Function
'
'
'
'Function PR_GetFirstPivotRow(ByRef vStorageObject As Variant, _
'    ByVal sReportName As String) As Long
'
'    Dim ls As ListStorage
'
'    Set ls = vStorageObject
'
'    PR_GetFirstPivotRow = ls.Xlookup(sReportName & "Sheet" & "PivotTableFirstRow", _
'    "[ReportName] & [DataType] & [Property]", "[Value]")
'
'End Function
