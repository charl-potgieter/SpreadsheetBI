Attribute VB_Name = "m800_UTILITIES_Sundry"
Option Explicit
Option Private Module

'------------------------------------------------------------------------------x
'    Requires reference to Microsoft Forms 2.0 object library
'------------------------------------------------------------------------------

'Below is utilisedto detect screen resolution
'https://www.mrexcel.com/board/threads/vba-to-find-screen-size-in-64bit-environment.1018797/
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1



Sub StandardEntry()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
End Sub


Sub StandardExit()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
End Sub



Function SheetLevelRangeNameExists(sht As Worksheet, ByVal sRangeName As String)
'Returns TRUE if sheet level scoped range name exists

    Dim sTest As String
    
    On Error Resume Next
    sTest = sht.Names(sRangeName).Name
    SheetLevelRangeNameExists = (Err.Number = 0)
    On Error GoTo 0


End Function


Function WorkbookLevelRangeNameExists(ByVal wkb As Workbook, ByVal sRangeName As String)
'Returns TRUE if sheet level scoped range name exists

    Dim sTest As String
    
    On Error Resume Next
    sTest = wkb.Names(sRangeName).Name
    WorkbookLevelRangeNameExists = (Err.Number = 0)
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

    Dim item As Variant
    
    ValueIsInStringArray = False
    For Each item In aArray
        If item = aValueToTest Then
            ValueIsInStringArray = True
            Exit Function
        End If
    Next item

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


Function UserFormComboBoxHasSelectedItems(ByRef combo As MSForms.ComboBox) As Boolean
    UserFormComboBoxHasSelectedItems = (combo.ListIndex <> -1)
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

Function GetCellColour(rng As Range, Optional formatType As Integer = 0) As Variant

'https://stackoverflow.com/questions/24132665/return-rgb-values-from-range-interior-color-or-any-other-color-property?rq=1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function            Color
'   Purpose             Determine the Background Color Of a Cell
'   @Param rng          Range to Determine Background Color of
'   @Param formatType   Default Value = 0
'                       0   Integer
'                       1   Hex
'                       2   RGB
'                       3   Excel Color Index
'   Usage               Color(A1)      -->   9507341
'                       Color(A1, 0)   -->   9507341
'                       Color(A1, 1)   -->   91120D
'                       Color(A1, 2)   -->   13, 18, 145
'                       Color(A1, 3)   -->   6
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Dim colorVal As Variant
    colorVal = Cells(rng.Row, rng.Column).Interior.Color
    Select Case formatType
        Case 1
            GetCellColour = Hex(colorVal)
        Case 2
            GetCellColour = (colorVal Mod 256) & ", " & ((colorVal \ 256) Mod 256) & ", " & (colorVal \ 65536)
        Case 3
            GetCellColour = Cells(rng.Row, rng.Column).Interior.ColorIndex
        Case Else
            GetCellColour = colorVal
    End Select
End Function




Function GetCellFontColour(rng As Range, Optional formatType As Integer = 0) As Variant

'https://stackoverflow.com/questions/24132665/return-rgb-values-from-range-interior-color-or-any-other-color-property?rq=1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function            Color
'   Purpose             Determine the Font Color Of a Cell
'   @Param rng          Range to Determine Background Color of
'   @Param formatType   Default Value = 0
'                       0   Integer
'                       1   Hex
'                       2   RGB
'                       3   Excel Color Index
'   Usage               Color(A1)      -->   9507341
'                       Color(A1, 0)   -->   9507341
'                       Color(A1, 1)   -->   91120D
'                       Color(A1, 2)   -->   13, 18, 145
'                       Color(A1, 3)   -->   6
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Dim colorVal As Variant
    colorVal = Cells(rng.Row, rng.Column).Font.Color
    Select Case formatType
        Case 1
            GetCellFontColour = Hex(colorVal)
        Case 2
            GetCellFontColour = (colorVal Mod 256) & ", " & ((colorVal \ 256) Mod 256) & ", " & (colorVal \ 65536)
        Case 3
            GetCellFontColour = Cells(rng.Row, rng.Column).Interior.ColorIndex
        Case Else
            GetCellFontColour = colorVal
    End Select
End Function





Function GetDataValidationFromRangeReference(ByVal rngSingleCell As Range) As Variant
'Returns a variant array of data validation items for rngSingleCell
'rngSingleCell must contain a list validation
    
    Dim bValidationListIsRange
    Dim rngValidationReference
    Dim sValidationFormula As String
    Dim SplitStringArray As Variant
    Dim ReturnValue() As Variant
    Dim i As Long
    
    On Error Resume Next
    sValidationFormula = rngSingleCell.Validation.Formula1
    If Err.Number <> 0 Then
        GetDataValidationFromRangeReference = Nothing
        Exit Function
    End If
    
    Set rngValidationReference = Range(Replace(sValidationFormula, "=", ""))
    bValidationListIsRange = (Err.Number = 0)
    On Error GoTo 0

    If bValidationListIsRange Then
        ReDim ReturnValue(0 To (rngValidationReference.Cells.Count - 1))
        For i = LBound(ReturnValue) To UBound(ReturnValue)
            ReturnValue(i) = rngValidationReference.Cells(i + 1)
        Next i
    Else
        SplitStringArray = Split(sValidationFormula, ",")
        ReDim ReturnValue(LBound(SplitStringArray) To UBound(SplitStringArray))
        For i = LBound(SplitStringArray) To UBound(SplitStringArray)
            ReturnValue(i) = SplitStringArray(i)
        Next i
    End If
    
    GetDataValidationFromRangeReference = ReturnValue

End Function



Function UniqueCollectionFromArray(ByVal arr) As Collection

    Dim col As Collection
    Dim i As Long
    
    Set col = New Collection

    On Error Resume Next
    For i = LBound(arr) To UBound(arr)
        col.Add item:=arr(i), key:=arr(i)
    Next i
    
    On Error GoTo 0


    Set UniqueCollectionFromArray = col
    Set col = Nothing

End Function


Sub BubbleSortSortCollection(ByRef col As Collection)

    Dim item As Variant
    Dim i As Long
    Dim TempItem As Variant
    Dim SwapMade As Boolean

    SwapMade = True
    Do Until SwapMade = False
        SwapMade = False
        For i = 1 To col.Count - 1
            If col(i) > col(i + 1) Then
                TempItem = col(i)
                col.Remove (i)
                col.Add item:=TempItem, key:=TempItem, After:=i
                SwapMade = True
            End If
        Next i
    Loop

End Sub


Function NumberOfSelectedItemsInListBox(ByVal lb As MSForms.ListBox) As Long

    Dim i As Long
    Dim NumberOfSelections As Long

    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) = True Then
            NumberOfSelections = NumberOfSelections + 1
        End If
    Next i
    NumberOfSelectedItemsInListBox = NumberOfSelections

End Function




Function ArrayOfListBoxSelections(ByVal lb As MSForms.ListBox)

    Dim i As Long
    Dim j As Long
    Dim ReturnArray() As Variant
    

    ReDim ReturnArray(0 To lb.ListCount - 1)
    j = 0
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) Then
            ReturnArray(j) = lb.List(i)
            j = j + 1
        End If
    Next i
    ReDim Preserve ReturnArray(0 To j - 1)
    ArrayOfListBoxSelections = ReturnArray

End Function



Function FirstListBoxSelection(ByVal lb As MSForms.ListBox)

    Dim i As Long
    Dim SelectionFound As Boolean
    
    i = 0
    SelectionFound = False
    Do While i <= (lb.ListCount - 1) And Not SelectionFound
        If lb.Selected(i) Then
            FirstListBoxSelection = lb.List(i)
            SelectionFound = True
        End If
        i = i + 1
    Loop
    
End Function


Function CleanTrim(ByVal str As String) As String
    CleanTrim = WorksheetFunction.Clean(Trim(str))
End Function


Function StringIsARangeReference(ByVal ReferenceStr As String) As Boolean

    Dim ReferenceAndEquals As String
    Dim ReferenceWithoutEquals As String
    
    If Left(ReferenceStr, 1) <> "=" Then
        ReferenceAndEquals = "=" & ReferenceStr
        ReferenceWithoutEquals = ReferenceStr
    Else
        ReferenceAndEquals = ReferenceStr
        ReferenceWithoutEquals = Right(ReferenceStr, Len(ReferenceStr) - 1)
    End If
    
    On Error Resume Next
    Select Case True
        Case Not (IsError(Application.Evaluate(ReferenceAndEquals)))
            StringIsARangeReference = True
        Case Not (IsError(Application.Evaluate(ReferenceAndEquals)(0)))
            StringIsARangeReference = True
        Case Else
            StringIsARangeReference = False
    End Select
    If Err.Number <> 0 Then StringIsARangeReference = False
    On Error GoTo 0

End Function


Function ConvertReferenceToStructuredReference(ByVal CallingCell As Range, ByVal RefStr As String) As String
'Converts a range reference to a structured table reference if it exists, otherwise leave reference unchanged
'CallingCell is the cell where the reference will be placed

    Dim TargetTable As ListObject
    Dim rngCell As Range
    Dim lo As ListObject
    Dim ListObjName As String
    Dim ReferenceRange As Range
    Dim ReferenceIsInsideSingleListObject As Boolean
    Dim i As Long
    Dim TableColumnNumberSelected As Integer
   
    On Error Resume Next
    Set ReferenceRange = Range(RefStr)
    If Err.Number <> 0 Then
        ConvertReferenceToStructuredReference = RefStr
        GoTo Exitpoint
    End If
    
    'Check if entire reference is inside a single listobject
    i = 1
    ReferenceIsInsideSingleListObject = True
    Do While i <= ReferenceRange.Cells.Count And ReferenceIsInsideSingleListObject
        Select Case True
            Case ReferenceRange.Cells(i).ListObject Is Nothing
                ReferenceIsInsideSingleListObject = False
            Case i = 1
                ListObjName = ReferenceRange.Cells(i).ListObject.Name
            Case Else
                ReferenceIsInsideSingleListObject = ReferenceIsInsideSingleListObject And _
                    ReferenceRange.Cells(i).ListObject.Name = ListObjName
        End Select
        i = i + 1
    Loop
    
    If Not ReferenceIsInsideSingleListObject Then
        ConvertReferenceToStructuredReference = RefStr
        GoTo Exitpoint
    End If
    
    Set lo = ReferenceRange.Cells(1).ListObject
    
    Select Case True
    
        'Entire table selected
        Case ReferenceRange.Address = lo.Range.Address
            ConvertReferenceToStructuredReference = lo.Name & "[#All]"
            
        'Entire databody range selected
        Case ReferenceRange.Address = lo.DataBodyRange.Address
            ConvertReferenceToStructuredReference = lo.Name
            
        'Entire columns databody range selected
        Case (ReferenceRange.Columns.Count = 1) And (ReferenceRange.Rows.Count = lo.DataBodyRange.Rows.Count)
            TableColumnNumberSelected = ReferenceRange.Column - lo.Range.Cells(1).Column + 1
            ConvertReferenceToStructuredReference = lo.Name & "[" & _
                lo.ListColumns(TableColumnNumberSelected).Name & "]"
                
        'Entire column selected including heading
        Case (ReferenceRange.Columns.Count = 1) And (ReferenceRange.Rows.Count = lo.Range.Rows.Count)
            TableColumnNumberSelected = ReferenceRange.Column - lo.Range.Cells(1).Column + 1
            ConvertReferenceToStructuredReference = lo.Name & "[[#All],[" & _
                lo.ListColumns(TableColumnNumberSelected).Name & "]]"
                
        'One cell in databodyrange selected adjacent to the cell allowing @ style referencing
        Case ReferenceRange.Cells.Count = 1 And (CallingCell.Row = ReferenceRange.Row)
            TableColumnNumberSelected = ReferenceRange.Column - lo.Range.Cells(1).Column + 1
            ConvertReferenceToStructuredReference = lo.Name & "[@" & _
                lo.ListColumns(TableColumnNumberSelected).Name & "]"
        
        'Single cell in header is selected
        Case ReferenceRange.Cells.Count = 1 And ReferenceRange.Row = lo.HeaderRowRange.Row
            TableColumnNumberSelected = ReferenceRange.Column - lo.Range.Cells(1).Column + 1
            ConvertReferenceToStructuredReference = lo.Name & "[[#Headers],[" & _
                lo.ListColumns(TableColumnNumberSelected).Name & "]]"
                
        'Entire table header is selected
        Case ReferenceRange.Address = lo.HeaderRowRange.Address
            ConvertReferenceToStructuredReference = lo.Name & "[#Headers]"
            
        'Not a structured reference
        Case Else
            ConvertReferenceToStructuredReference = RefStr
            
    End Select
    

Exitpoint:
    Set TargetTable = Nothing
    Set lo = Nothing
    Set ReferenceRange = Nothing

End Function


Function ListObjHasQueryTable(ByVal lo As ListObject) As Boolean

    Dim i As Integer
    
    On Error Resume Next
    i = lo.QueryTable.CommandType
    ListObjHasQueryTable = (Err.Number = 0)
    On Error GoTo 0

End Function

