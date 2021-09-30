Attribute VB_Name = "m020_DATA_ACCESS_SundryStorage"
Option Explicit
Option Private Module


Sub DeleteSundryStorageItems(ByVal Item As String)

    Dim Storage As ListStorage
    
    Set Storage = New ListStorage
    Storage.AssignStorage ThisWorkbook, "SundryStorage"
    
    Storage.Filter "[Item] <> """ & Item & """"
    Storage.ReplaceDataWithFilteredData

End Sub



Function GetSundryStorageItem(ByVal Item As String) As Variant

    Dim Storage As ListStorage
    
    Set Storage = New ListStorage
    Storage.AssignStorage ThisWorkbook, "SundryStorage"
    GetSundryStorageItem = Storage.Xlookup(Item, "[Item]", "[Value]")

End Function

