Attribute VB_Name = "m020_DATA_ACCESS_Developer"
'@Folder "SpreadsheetBI"
Option Explicit
Option Private Module

Function CreateListObjFieldStorage(ByVal SourceFilePath As String, _
    ByVal TargetStorage) As ListStorage
    
    Dim Storage As ListStorage

    CreatePipeDelimitedPowerQuery TargetStorage, _
        SourceFilePath, _
        "ListObjectFields"

    Set Storage = New ListStorage
    Storage.CreateStorageFromPowerQuery TargetStorage, "ListObjectFields", "ListObjectFields"
    
    Set CreateListObjFieldStorage = Storage
    
End Function



Function CreateListObjFieldValuesStorage(ByVal SourceFilePath As String, _
    ByVal TargetStorage) As ListStorage
    
    Dim Storage As ListStorage

    CreatePipeDelimitedPowerQuery TargetStorage, _
        SourceFilePath, _
        "ListObjectFieldValues"

    Set Storage = New ListStorage
    Storage.CreateStorageFromPowerQuery TargetStorage, "ListObjectFieldValues", _
        "ListObjectFieldValues"
    
    Set CreateListObjFieldValuesStorage = Storage
    
End Function



Function CreateListObjFieldFormatsStorage(ByVal SourceFilePath As String, _
    ByVal TargetStorage) As ListStorage
    
    Dim Storage As ListStorage

    CreatePipeDelimitedPowerQuery TargetStorage, _
        SourceFilePath, _
        "ListObjectFieldFormats"

    Set Storage = New ListStorage
    Storage.CreateStorageFromPowerQuery TargetStorage, "ListObjectFieldFormats", _
        "ListObjectFieldFormats"
    
    Set CreateListObjFieldFormatsStorage = Storage
    
End Function


Function CreateOtherStorage(ByVal SourceFilePath As String, _
    ByVal TargetStorage) As ListStorage
    
    Dim Storage As ListStorage

    CreatePipeDelimitedPowerQuery TargetStorage, _
        SourceFilePath, _
        "OtherData"

    Set Storage = New ListStorage
    Storage.CreateStorageFromPowerQuery TargetStorage, "OtherData", "OtherData"
    
    Set CreateOtherStorage = Storage
    
End Function


Function GetSheetNames(ByVal ListgObjFieldStorage As Variant) As Variant

    Dim Storage As ListStorage
    
    Set Storage = ListgObjFieldStorage
    If Storage.NumberOfRecords <> 0 Then
        GetSheetNames = Storage.ItemsInField(sFieldName:="SheetName", bUnique:=True)
    End If
    Set Storage = Nothing

End Function



Private Sub CreatePipeDelimitedPowerQuery(ByVal wkb As Workbook, _
    ByVal SourceDelimitedFilePath As String, _
    ByVal QueryName As String)

    Dim QueryString As String
    
    QueryString = _
        "let" & vbCr & _
        "    Source = Csv.Document(File.Contents(""" & _
        SourceDelimitedFilePath & """" & _
        "),[Delimiter=""|"", Encoding=1252, QuoteStyle=QuoteStyle.None])," & vbCr & _
        "   PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])" & vbCr & _
        "in " & vbCr & _
        "   PromotedHeaders"
    
    wkb.Queries.Add QueryName, QueryString

End Sub



Function GetCreatorFileName(ByVal StorageOther As Variant) As String

    Dim Storage As ListStorage
    Set Storage = StorageOther
    GetCreatorFileName = Storage.Xlookup("FileName", "[Item]", "[Value]")
    Set Storage = Nothing

End Function



Function GetListObjName(ByVal StorageListObjFields As Variant, _
    ByVal SheetName As String) As String

    Dim Storage As ListStorage
    Set Storage = StorageListObjFields
    GetListObjName = Storage.Xlookup(SheetName, "[SheetName]", "[ListObjectName]")
    Set Storage = Nothing

End Function


Function GetListObjHeaders(ByVal StorageListObjFields As Variant, _
    ByVal SheetName As String) As Variant

    Dim Storage As ListStorage
    Set Storage = StorageListObjFields
    
    Storage.Filter "[SheetName] = """ & SheetName & """"
    GetListObjHeaders = Storage.ItemsInField(sFieldName:="ListObjectHeader", bFiltered:=True)

End Function


Function GetHeaderHasFormula(ByVal StorageListObjFields As Variant, _
    ByVal SheetName As String, ByVal Header As String) As Boolean

    Dim Storage As ListStorage
    
    Set Storage = StorageListObjFields
    GetHeaderHasFormula = Storage.Xlookup(SheetName & Header, _
        "[SheetName] & [ListObjectHeader]", "[IsFormula]")
    
    Set Storage = Nothing

End Function



Function GetColumnFormula(ByVal StorageListObjFields As Variant, _
    ByVal SheetName As String, ByVal Header As String) As String

    Dim Storage As ListStorage
    
    Set Storage = StorageListObjFields
    GetColumnFormula = Storage.Xlookup(SheetName & Header, _
        "[SheetName] & [ListObjectHeader]", "[Formula]")
    
    Set Storage = Nothing

End Function


Function GetColumnNumberFormat(ByVal StorageListObjFieldFormats As Variant, _
    ByVal SheetName As String, ByVal Header As String) As String

    Dim Storage As ListStorage
    
    Set Storage = StorageListObjFieldFormats
    GetColumnNumberFormat = Storage.Xlookup(SheetName & Header, _
        "[SheetName] & [ListObjectHeader]", "[NumberFormat]")
    
    Set Storage = Nothing

End Function


Function GetColumnFontColour(ByVal StorageListObjFieldFormats As Variant, _
    ByVal SheetName As String, ByVal Header As String) As String

    Dim Storage As ListStorage
    
    Set Storage = StorageListObjFieldFormats
    GetColumnFontColour = Storage.Xlookup(SheetName & Header, _
        "[SheetName] & [ListObjectHeader]", "[FontColour]")
    
    Set Storage = Nothing

End Function


Function GetTableValues(ByVal StorageListObjFieldValues As Variant, _
    ByVal SheetName As String) As Dictionary()

    Dim Storage As ListStorage
    Dim dict() As Dictionary
    Dim CurrentDictArrayIndex As Long
    Dim CurrentDataRow As Long
    Dim Header As String
    Dim Value As Variant
    Dim NumberOfDataRows As Long
    Dim MaxAllowAlowableDataRows As Long

    MaxAllowAlowableDataRows = ActiveSheet.Rows.Count
    ReDim dict(0 To MaxAllowAlowableDataRows - 1)
    Set Storage = StorageListObjFieldValues

    Storage.Filter "[SheetName] = """ & SheetName & """"
    CurrentDictArrayIndex = 0
    CurrentDataRow = 1
    NumberOfDataRows = Storage.NumberOfRecords(bFiltered:=True)
    
    
    If NumberOfDataRows <> 0 Then
        Set dict(CurrentDictArrayIndex) = New Dictionary
        Do While CurrentDataRow <= NumberOfDataRows
            Header = Storage.FieldItemByIndex( _
                sFieldName:="ListObjectHeader", _
                i:=CurrentDataRow, _
                bFiltered:=True)
            Value = Storage.FieldItemByIndex( _
                sFieldName:="Value", _
                i:=CurrentDataRow, _
                bFiltered:=True)
            If dict(CurrentDictArrayIndex).Exists(Header) Then
                CurrentDictArrayIndex = CurrentDictArrayIndex + 1
                Set dict(CurrentDictArrayIndex) = New Dictionary
            End If
            dict(CurrentDictArrayIndex).Add Header, Value
            CurrentDataRow = CurrentDataRow + 1
        Loop
    End If
    ReDim Preserve dict(0 To CurrentDictArrayIndex)

    Set Storage = Nothing
    GetTableValues = dict

End Function


Sub DeleteStorage(ByRef Storage)

    Storage.Delete

End Sub


Function StorageIsEmpty(ByVal Storage) As Boolean

    StorageIsEmpty = Storage.NumberOfRecords = 0

End Function


