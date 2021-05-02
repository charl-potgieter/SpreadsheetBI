Attribute VB_Name = "m040_SundryUtilities"
Option Explicit
Option Private Module

'Below is utilisedto detect screen resolution
'https://www.mrexcel.com/board/threads/vba-to-find-screen-size-in-64bit-environment.1018797/
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1



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





Function WorkbookIsOpen(ByVal sWbkName As String) As Boolean
'Checks if workbook is open based on filename including extension

    WorkbookIsOpen = False
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


Sub CommaSeperatedDelimit(ByVal sInput As String, ByRef aDelimited() As String)
    
    Dim i As Double

    aDelimited = Split(sInput, ",")
    
    'Trim any whitespace before or after commas
    For i = LBound(aDelimited) To UBound(aDelimited)
        aDelimited(i) = Trim(aDelimited(i))
    Next i

End Sub



Function ValueIsInStringArray(ByVal aValueToTest As Variant, ByRef aArray() As String) As Boolean

    Dim Item As Variant
    
    ValueIsInStringArray = False
    For Each Item In aArray
        If Item = aValueToTest Then
            ValueIsInStringArray = True
            Exit Function
        End If
    Next Item

End Function


Function UserFormListBoxHasSelectedItems(ByRef lb As MSForms.ListBox) As Boolean
    Dim i As Long
    UserFormListBoxHasSelectedItems = False
    i = 0
    Do While Not UserFormListBoxHasSelectedItems And i <= lb.ListCount - 1
        UserFormListBoxHasSelectedItems = lb.Selected(i)
        i = i + 1
    Loop
End Function


Function UserFormListBoxSelectedArray(ByVal lb As MSForms.ListBox) As String()

    Dim ReturnArray() As String
    Dim i As Long
    Dim j As Long

    'Leave array empty if nothing selected
    If Not UserFormListBoxHasSelectedItems(lb) Then Exit Function
    j = 0
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) = True Then
            ReDim Preserve ReturnArray(j)
            ReturnArray(j) = lb.List(i)
            j = j + 1
        End If
    Next i
    UserFormListBoxSelectedArray = ReturnArray

End Function



Public Function ScreenDimensionWidth() As Long
'https://www.mrexcel.com/board/threads/vba-to-find-screen-size-in-64bit-environment.1018797/
   
    'Note declaration function at top of module
    ScreenDimensionWidth = GetSystemMetrics(SM_CXSCREEN)
   
End Function


Public Function ScreenDimensionHeight() As Long
'https://www.mrexcel.com/board/threads/vba-to-find-screen-size-in-64bit-environment.1018797/
   
    'Note declaration function at top of module
    ScreenDimensionHeight = GetSystemMetrics(SM_CYSCREEN)
   
End Function
