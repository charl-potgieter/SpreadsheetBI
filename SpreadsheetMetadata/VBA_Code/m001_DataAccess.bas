Attribute VB_Name = "m001_DataAccess"
Option Explicit
Option Private Module


Function GetSavedReportSheetFormat() As Dictionary

    Dim i As Integer
    Dim Storage As zLIB_ListStorage
    Dim dict As Dictionary
    
    Set dict = New Dictionary
    Set Storage = New zLIB_ListStorage
    
    Storage.AssignStorage ThisWorkbook, "ReportSheetFormat"
    For i = 1 To Storage.NumberOfRecords
        dict.Add _
            key:=Storage.ListObj.ListColumns("Item").DataBodyRange.Cells(i).Value, _
            item:=Storage.ListObj.ListColumns("Value").DataBodyRange.Cells(i).Value
    Next i

    Set GetSavedReportSheetFormat = dict

End Function



