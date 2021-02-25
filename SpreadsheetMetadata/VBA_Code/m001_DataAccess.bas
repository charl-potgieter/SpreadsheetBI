Attribute VB_Name = "m001_DataAccess"
Option Explicit
Option Private Module

Const csPowerReportStorageName As String = "ReportSheetProperties"

Sub SetupOrAssignPowerReportStorage()
'Assigns storage to liststorage object if it exists, otthwerise creates new storage

    Dim ls As ListStorage
    Dim bStorageIsAssigned
    Dim sHeaders(4) As String
    
    Set ls = New ListStorage
    bStorageIsAssigned = ls.AssignStorage(ActiveWorkbook, csPowerReportStorageName)

    If Not bStorageIsAssigned Then
        'Storage not assigned, need to create
        sHeaders(0) = "SheetName"
        sHeaders(1) = "Name"
        sHeaders(2) = "DataType"
        sHeaders(3) = "Property"
        sHeaders(4) = "Value"
        ls.CreateStorage ActiveWorkbook, csPowerReportStorageName, sHeaders
    End If


End Sub



Sub DeleteExistingReportData(ByVal sSheetName As String)
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


Sub WritePowerReportRecord( _
    ByVal sSheetName As String, _
    ByVal sName As String, _
    ByVal sDataType As String, _
    ByVal sProperty As String, _
    ByVal sValue As String)
    
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
    ls.InsertFromDictionary Dict
    
End Sub
    
Function GetReportHeadingNameBasedOnSheetName(ByVal sSheetName As String) As String

    Dim ls As ListStorage
    Const csStorageName As String = "ReportSheetProperries"
        
    Set ls = New ListStorage
    ls.AssignStorage ActiveWorkbook, csPowerReportStorageName
    
    GetReportHeadingNameBasedOnSheetName = ls.Xlookup( _
        LookupValue:=sSheetName & "SheetDataType" & "SheetHeading", _
        sLookupArray:="[SheetName] & [DataType] & [Property]", _
        sReturnArray:="[Value]")
        
End Function


Function GetReportCategoryNameBasedOnSheetName(ByVal sSheetName As String) As String

    Dim ls As ListStorage
    Const csStorageName As String = "ReportSheetProperries"
        
    Set ls = New ListStorage
    ls.AssignStorage ActiveWorkbook, csPowerReportStorageName
    
    GetReportCategoryNameBasedOnSheetName = ls.Xlookup( _
        LookupValue:=sSheetName & "SheetDataType" & "SheetCategory", _
        sLookupArray:="[SheetName] & [DataType] & [Property]", _
        sReturnArray:="[Value]")
        
End Function

