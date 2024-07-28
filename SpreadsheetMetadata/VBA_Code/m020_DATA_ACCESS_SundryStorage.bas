Attribute VB_Name = "m020_DATA_ACCESS_SundryStorage"
Option Explicit
Option Private Module

Sub UpdateSundryStorageValueForGivenItem(ByVal item As String, ByVal Value As Variant)

    Dim Storage As ListStorage
    Dim dict As Dictionary
    
    'Delete existing data record
    Set Storage = New ListStorage
    Storage.AssignStorage ThisWorkbook, "SundryStorage"
    Storage.Filter "[Item] <> """ & item & """"
    Storage.ReplaceDataWithFilteredData
    
    'Insert new data record
    Set dict = New Dictionary
    dict.Add "Item", item
    dict.Add "Value", Value
    Storage.InsertFromDictionary dict

End Sub




Sub DeleteSundryStorageByItemValue(ByVal item As String)

    Dim Storage As ListStorage
    
    Set Storage = New ListStorage
    Storage.AssignStorage ThisWorkbook, "SundryStorage"
    
    Storage.Filter "[Item] <> """ & item & """"
    Storage.ReplaceDataWithFilteredData
    Set Storage = Nothing

End Sub


Function GetSundryStorageItem(ByVal item As String) As Variant

    Dim Storage As ListStorage

    Set Storage = New ListStorage
    Storage.AssignStorage ThisWorkbook, "SundryStorage"
    GetSundryStorageItem = Storage.Xlookup(item, "[Item]", "[Value]")
    Set Storage = Nothing

End Function


