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
            item:=Storage.ListObj.ListColumns("Value").DataBodyRange.Cells(i).Value
    Next i

    Set GetSavedReportSheetFormat = dict

End Function



