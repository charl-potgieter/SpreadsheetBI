Attribute VB_Name = "m040_SundryUtilities"
Option Explicit

Function SheetLevelRangeNameExists(sht As Worksheet, ByRef sRangeName As String)
'Returns TRUE if sheet level scoped range name exists

    Dim sTest As String
    
    On Error Resume Next
    sTest = sht.Names(sRangeName).Name
    SheetLevelRangeNameExists = (Err.Number = 0)
    On Error GoTo 0


End Function



Function SheetExists(wkb As Workbook, sSheetName As String) As Boolean

    SheetExists = False
    On Error Resume Next
    SheetExists = Len(wkb.Sheets(sSheetName).Name) > 0
    On Error GoTo 0
    
End Function



Function ArrayIsInitialised(arr) As Boolean

    Dim value
    
    On Error Resume Next
    value = arr(0)
    ArrayIsInitialised = (Err.Number = 0)
    On Error GoTo 0

End Function

