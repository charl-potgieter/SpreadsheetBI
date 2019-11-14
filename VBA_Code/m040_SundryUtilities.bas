Attribute VB_Name = "m040_SundryUtilities"
Option Explicit
Option Private Module


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


Function WorkbookIsOpen(ByVal sWbkName As String) As Boolean

    On Error Resume Next
    WorkbookIsOpen = Len(Workbooks(sWbkName).Name) <> 0
    On Error GoTo 0

End Function


Function ActiveCellIsInPivotTable() As Boolean

    Dim pvt As PivotTable
    
    On Error Resume Next
    Set pvt = ActiveCell.PivotTable
    ActiveCellIsInPivotTable = Len(pvt.Name) <> 0
    On Error GoTo 0
    
End Function


Function QueryExists(ByVal sQryName As String, Optional wkb As Workbook) As Boolean

    If wkb Is Nothing Then
        Set wkb = ActiveWorkbook
    End If
    
    On Error Resume Next
    QueryExists = CBool(Len(wkb.Queries(sQryName).Name))
    On Error GoTo 0
    

End Function



Function TableExistsInWorkbook(ByVal sTableName As String, Optional wkb As Workbook) As Boolean

    Dim lo As ListObject
    Dim sht As Worksheet
    Dim bTableFound As Boolean
    
    If wkb Is Nothing Then Set wkb = ActiveWorkbook
    bTableFound = False
    
    For Each sht In wkb.Worksheets
        For Each lo In sht.ListObjects
            If lo.Name = sTableName Then bTableFound = True
        Next lo
    Next sht
    
    TableExistsInWorkbook = bTableFound
    
End Function




Function TableExistsInSheet(ByVal sTableName As String, Optional sht As Worksheet) As Boolean

    Dim lo As ListObject
    Dim bTableFound As Boolean
    
    If sht Is Nothing Then Set sht = ActiveSheet
    bTableFound = False
    
    For Each lo In sht.ListObjects
        If lo.Name = sTableName Then bTableFound = True
    Next lo
    
    TableExistsInSheet = bTableFound

End Function






















