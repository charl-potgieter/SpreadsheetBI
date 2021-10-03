Attribute VB_Name = "m020_DATA_ACCESS_SundryStorage"
Option Explicit
Option Private Module

Sub UpdateSundryStorageValueForGivenItem(ByVal Item As String, ByVal Value As Variant)

    Dim Storage As ListStorage
    Dim dict As Dictionary
    
    'Delete existing data record
    Set Storage = New ListStorage
    Storage.AssignStorage ThisWorkbook, "SundryStorage"
    Storage.Filter "[Item] <> """ & Item & """"
    Storage.ReplaceDataWithFilteredData
    
    'Insert new data record
    Set dict = New Dictionary
    dict.Add "Item", Item
    dict.Add "Value", Value
    Storage.InsertFromDictionary dict

End Sub




Sub DeleteSundryStorageByItemValue(ByVal Item As String)

    Dim Storage As ListStorage
    
    Set Storage = New ListStorage
    Storage.AssignStorage ThisWorkbook, "SundryStorage"
    
    Storage.Filter "[Item] <> """ & Item & """"
    Storage.ReplaceDataWithFilteredData
    Set Storage = Nothing

End Sub


Function GetSundryStorageItem(ByVal Item As String) As Variant

    Dim Storage As ListStorage
    
    Set Storage = New ListStorage
    Storage.AssignStorage ThisWorkbook, "SundryStorage"
    GetSundryStorageItem = Storage.Xlookup(Item, "[Item]", "[Value]")
    Set Storage = Nothing

End Function


