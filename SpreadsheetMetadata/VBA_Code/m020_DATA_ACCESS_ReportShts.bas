Attribute VB_Name = "m020_DATA_ACCESS_ReportShts"
Option Explicit
Option Private Module


Function GetSavedReportSheetFormat() As Dictionary

    Dim i As Integer
    Dim Storage As ListStorage
    Dim dict As Dictionary

    Set dict = New Dictionary
    Set Storage = New ListStorage

    Storage.AssignStorage ThisWorkbook, "ReportSheetFormat"
    For i = 1 To Storage.NumberOfRecords
        dict.Add _
            key:=Storage.ListObj.ListColumns("Item").DataBodyRange.Cells(i).Value, _
            Item:=Storage.ListObj.ListColumns("Value").DataBodyRange.Cells(i).Value
    Next i

    Set GetSavedReportSheetFormat = dict

End Function



Function GetReportSheetFormatItem(ByVal Item As String) As Variant

    Dim Storage As ListStorage
    
    Set Storage = New ListStorage
    Storage.AssignStorage ThisWorkbook, "ReportSheetFormat"
    GetReportSheetFormatItem = Storage.Xlookup(Item, "[Item]", "[Value]")

End Function



Sub WriteReportSheetFormat(ByRef ReportSheetFormatDict() As Dictionary)

    Dim i As Integer
    Dim Storage As ListStorage
    
    Set Storage = New ListStorage
    Storage.AssignStorage ThisWorkbook, "ReportSheetFormat"
    Storage.ClearData
    For i = LBound(ReportSheetFormatDict) To UBound(ReportSheetFormatDict)
        Storage.InsertFromDictionary ReportSheetFormatDict(i)
    Next i

End Sub


