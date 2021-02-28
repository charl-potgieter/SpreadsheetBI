Attribute VB_Name = "m001_DataAccess"
Option Explicit
Option Private Module

Const csPowerReportStorageName As String = "ReportSheetProperties"

Sub PR_SetupOrAssignStorage()
'Assigns storage to liststorage object if it exists, otthwerise creates new storage

    Dim ls As ListStorage
    Dim bStorageIsAssigned
    Dim sHeaders(5) As String
    
    Set ls = New ListStorage
    bStorageIsAssigned = ls.AssignStorage(ActiveWorkbook, csPowerReportStorageName)

    If Not bStorageIsAssigned Then
        'Storage not assigned, need to create
        sHeaders(0) = "SheetName"
        sHeaders(1) = "Name"
        sHeaders(2) = "DataType"
        sHeaders(3) = "Property"
        sHeaders(4) = "Value"
        sHeaders(5) = "CubeFieldPosition"
        ls.CreateStorage ActiveWorkbook, csPowerReportStorageName, sHeaders
    End If


End Sub



Sub PR_DeleteExistingData(ByVal sSheetName As String)
'If there is any existing data in Listorage with same sSheetname as report then this is deleted
    
    Dim sFilterString As String
    Dim ls As ListStorage
    
    Set ls = New ListStorage
    ls.AssignStorage ActiveWorkbook, csPowerReportStorageName
    
    'Create Filter excluding Report name and then replace data with filter
    sFilterString = "[SheetName] <> """ & sSheetName & """"
    ls.Filter sFilterString
    ls.ReplaceDataWithFilteredData
        
End Sub


Sub PR_WriteRecords( _
    ByVal sSheetName As String, _
    ByVal sName As String, _
    ByVal sDataType As String, _
    ByVal sProperty As String, _
    ByVal sValue As String, _
    ByVal lPosition As Variant)
    
'Writes a single property of PowerReport to list storage
    
    Dim Dict As Dictionary
    Dim ls As ListStorage
    
    Set ls = New ListStorage
    ls.AssignStorage ActiveWorkbook, csPowerReportStorageName

    Set Dict = New Dictionary
    Dict.Add "SheetName", sSheetName
    Dict.Add "Name", sName
    Dict.Add "DataType", sDataType
    Dict.Add "Property", sProperty
    Dict.Add "Value", sValue
    Dict.Add "CubeFieldPosition", lPosition
    ls.InsertFromDictionary Dict
    
End Sub
    
Function PR_GetHeadingNameBasedOnSheetName(ByVal sSheetName As String) As String

    Dim ls As ListStorage
    Const csStorageName As String = "ReportSheetProperries"
        
    Set ls = New ListStorage
    ls.AssignStorage ActiveWorkbook, csPowerReportStorageName
    
    PR_GetHeadingNameBasedOnSheetName = ls.Xlookup( _
        LookupValue:=sSheetName & "SheetDataType" & "SheetHeading", _
        sLookupArray:="[SheetName] & [DataType] & [Property]", _
        sReturnArray:="[Value]")
        
End Function


Function PR_GetCategoryNameBasedOnSheetName(ByVal sSheetName As String) As String

    Dim ls As ListStorage
    Const csStorageName As String = "ReportSheetProperries"
        
    Set ls = New ListStorage
    ls.AssignStorage ActiveWorkbook, csPowerReportStorageName
    
    PR_GetCategoryNameBasedOnSheetName = ls.Xlookup( _
        LookupValue:=sSheetName & "SheetDataType" & "SheetCategory", _
        sLookupArray:="[SheetName] & [DataType] & [Property]", _
        sReturnArray:="[Value]")
        
End Function

Function PR_GetPivotTableProperties(ByVal sSheetName As String, _
    ByRef StorageRecords() As TypePowerReportStorageRecord)
'Returns  Pivot Table properties in the Storage Array.

    
    Dim ls As ListStorage
    Dim i As Long
    Dim sFilterStr As String

    Set ls = New ListStorage
    ls.AssignStorage ActiveWorkbook, csPowerReportStorageName

    sFilterStr = "([SheetName]=""" & sSheetName & """) * " & _
        "([DataType] = ""PivotTableDataType"")"

    ls.Filter sFilterStr

    ReDim StorageRecords(ls.NumberOfRecords(bFiltered:=True) - 1)
    
    For i = LBound(StorageRecords) To UBound(StorageRecords)
        StorageRecords(i).SheetName = ls.FieldItemByIndex("SheetName", i + 1, True)
        StorageRecords(i).Name = ls.FieldItemByIndex("Name", i + 1, True)
        StorageRecords(i).DataType = ls.FieldItemByIndex("DataType", i + 1, True)
        StorageRecords(i).Property = ls.FieldItemByIndex("Property", i + 1, True)
        StorageRecords(i).Value = ls.FieldItemByIndex("Value", i + 1, True)
    Next i

End Function


Function PR_GetPivotCubeFieldDataOrientationSortedByCubeFieldPosition(ByVal sSheetName As String, _
    ByRef StorageRecords() As TypePowerReportStorageRecord)
'Returns storage orientation property of pivot cube fields in the
'Storage Array.   Sorted by CubefieldPosition to ensure correct position when pivot table is created
'For some reason it doesn't work setting the position property directly in VBA

    
    Dim ls As ListStorage
    Dim i As Long
    Dim sFilterStr As String

    Set ls = New ListStorage
    ls.AssignStorage ActiveWorkbook, csPowerReportStorageName
    sFilterStr = "([SheetName]=""" & sSheetName & """) * " & _
        "([DataType] = ""PivotCubeFieldDataType"") * " & _
        "([Property] = ""Orientation"")"

    ls.Filter sFilterStr, True, "[CubeFieldPosition]", lsAsc

    ReDim StorageRecords(ls.NumberOfRecords(bFiltered:=True) - 1)
    
    For i = LBound(StorageRecords) To UBound(StorageRecords)
        StorageRecords(i).SheetName = ls.FieldItemByIndex("SheetName", i + 1, True)
        StorageRecords(i).Name = ls.FieldItemByIndex("Name", i + 1, True)
        StorageRecords(i).DataType = ls.FieldItemByIndex("DataType", i + 1, True)
        StorageRecords(i).Property = ls.FieldItemByIndex("Property", i + 1, True)
        StorageRecords(i).Value = ls.FieldItemByIndex("Value", i + 1, True)
    Next i

End Function


Function PR_GetPivotCubeFieldDataPropertiesExOrientation(ByVal sSheetName As String, _
    ByRef StorageRecords() As TypePowerReportStorageRecord)
'Returns storage properties ex-orientation property of pivot cube fields to the
'Storage Array

    
    Dim ls As ListStorage
    Dim i As Long

    Set ls = New ListStorage
    ls.AssignStorage ActiveWorkbook, csPowerReportStorageName
    ls.Filter "([SheetName]=""" & sSheetName & """) * " & _
        "([DataType] = ""PivotCubeFieldDataType"") * " & _
        "([Property] <> ""Orientation"")"

    ReDim StorageRecords(ls.NumberOfRecords(bFiltered:=True) - 1)
    
    For i = LBound(StorageRecords) To UBound(StorageRecords)
        StorageRecords(i).SheetName = ls.FieldItemByIndex("SheetName", i + 1, True)
        StorageRecords(i).Name = ls.FieldItemByIndex("Name", i + 1, True)
        StorageRecords(i).DataType = ls.FieldItemByIndex("DataType", i + 1, True)
        StorageRecords(i).Property = ls.FieldItemByIndex("Property", i + 1, True)
        StorageRecords(i).Value = ls.FieldItemByIndex("Value", i + 1, True)
    Next i

End Function

Function PR_GetPivotFieldDataSubtotalProperty(ByVal sSheetName As String, _
    ByRef StorageRecords() As TypePowerReportStorageRecord)
'Returns storage properties of pivot field subtotal to the Storage Array
'Subtotal handled seperately as it has flow on effects and slightly different in that property
'needs to be indexed

    
    Dim ls As ListStorage
    Dim i As Long

    Set ls = New ListStorage
    ls.AssignStorage ActiveWorkbook, csPowerReportStorageName
    ls.Filter "([SheetName]=""" & sSheetName & """) * " & _
        "([DataType] = ""PivotFieldDataType"") * " & _
        "([Property] = ""Subtotals"")"

    ReDim StorageRecords(ls.NumberOfRecords(bFiltered:=True) - 1)
    
    For i = LBound(StorageRecords) To UBound(StorageRecords)
        StorageRecords(i).SheetName = ls.FieldItemByIndex("SheetName", i + 1, True)
        StorageRecords(i).Name = ls.FieldItemByIndex("Name", i + 1, True)
        StorageRecords(i).DataType = ls.FieldItemByIndex("DataType", i + 1, True)
        StorageRecords(i).Property = ls.FieldItemByIndex("Property", i + 1, True)
        StorageRecords(i).Value = ls.FieldItemByIndex("Value", i + 1, True)
    Next i

End Function


Function PR_GetPivotFieldDataPropertiesExSubtotals(ByVal sSheetName As String, _
    ByRef StorageRecords() As TypePowerReportStorageRecord)
'Returns storage properties of pivot fields to the Storage Array

    
    Dim ls As ListStorage
    Dim i As Long

    Set ls = New ListStorage
    ls.AssignStorage ActiveWorkbook, csPowerReportStorageName
    ls.Filter "([SheetName]=""" & sSheetName & """) * " & _
        "([DataType] = ""PivotFieldDataType"") * " & _
        "([Property] <> ""Subtotals"")"

    ReDim StorageRecords(ls.NumberOfRecords(bFiltered:=True) - 1)
    
    For i = LBound(StorageRecords) To UBound(StorageRecords)
        StorageRecords(i).SheetName = ls.FieldItemByIndex("SheetName", i + 1, True)
        StorageRecords(i).Name = ls.FieldItemByIndex("Name", i + 1, True)
        StorageRecords(i).DataType = ls.FieldItemByIndex("DataType", i + 1, True)
        StorageRecords(i).Property = ls.FieldItemByIndex("Property", i + 1, True)
        StorageRecords(i).Value = ls.FieldItemByIndex("Value", i + 1, True)
    Next i

End Function

