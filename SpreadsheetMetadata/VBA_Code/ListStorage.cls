VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder "Storage"


'Use "This" declaration as an easy way to get intellisense to the classes private variables
'https://rubberduckvba.wordpress.com/2020/02/27/vba-classes-gateway-to-solid/
Private Type TypeListStorage
    wkb As Workbook
    Sheet As Worksheet
    ListObj As ListObject
    FilteredListObj As ListObject
End Type
Private this As TypeListStorage

Public Enum lsSortOrder  'ls = ListStorage
    lsAsc = 1
    lsDesc = -1
End Enum


Public Function CreateStorage(ByRef wkb As Workbook, ByVal sName As String, Headings() As String) As Boolean
'Does nothing and returns FALSE if sheet or listobject already exist
'Creates a new sheet named "sName" and listobject named "tbl_sName" in wkb and returns TRUE
    
    If StorageAlreadyExists(wkb, sName) Then
        CreateStorage = False
    Else
        CreateStorage = True
        
        Set this.wkb = wkb
        Set this.Sheet = CreateListStorageSheet(wkb, sName)
        FormatListStorageSheet this.Sheet
        
        Set this.ListObj = CreateListStorageListObject(this.Sheet, Headings)
        FormatStorageListObject this.ListObj
        
        Set this.FilteredListObj = CreateFilteredListStorageListObject(this.Sheet, Headings)
        FormatStorageListObject this.FilteredListObj
    End If

End Function


Public Function AssignStorage(ByRef wkb As Workbook, ByVal sName As String) As Boolean
'Assigns the storage object to a specific sheet and listobject
    
    
    On Error Resume Next
    
    Set this.wkb = wkb
    Set this.Sheet = wkb.Worksheets(sName)
    Set this.ListObj = this.Sheet.ListObjects("tbl_" & sName)
    Set this.FilteredListObj = this.Sheet.ListObjects("tbl_" & sName & "_Filtered")
    
    this.ListObj.ShowAutoFilter = True
    this.ListObj.AutoFilter.ShowAllData
        
    AssignStorage = (Err.Number = 0)
    
    On Error GoTo 0

End Function


Public Property Get Name() As String
    Name = this.Sheet.Name
End Property


Public Property Get NumberOfRecords(Optional bFiltered As Boolean = False) As Long

    On Error Resume Next

    NumberOfRecords = SelectedListObject(bFiltered).DataBodyRange.Rows.Count
    
    If Err.Number <> 0 Then
        'There is no databodyrange
        NumberOfRecords = 0
    End If
    On Error GoTo 0
    
End Property


Public Function ItemsInField(ByVal sFieldName As String, _
    Optional bIgnoreBlanks As Boolean = True, _
    Optional bUnique As Boolean = False, _
    Optional bSorted As Boolean = False, _
    Optional SortOrder As lsSortOrder = lsAsc, _
    Optional bFiltered As Boolean = False) As Variant
'Returns a list of items in sFieldname as a one dimensional variant array with base 1
'If empty or error returns null

    Dim sEvalStr As String
    Dim sTableName As String
    
    sTableName = SelectedListObject(bFiltered).Name
    
    If bIgnoreBlanks Then
        sEvalStr = "FILTER(<<Table>>[<<Field>>], <<Table>>[<<Field>>]<>"""")"
    Else
        sEvalStr = "SUBSTITUTE(<<Table>>[<<Field>>], """", """")"
    End If
    
    sEvalStr = Replace(sEvalStr, "<<Table>>", sTableName)
    sEvalStr = Replace(sEvalStr, "<<Field>>", sFieldName)
    
    If bUnique Then
        sEvalStr = "UNIQUE(" & sEvalStr & ")"
    End If
    If bSorted And SortOrder = lsAsc Then
        sEvalStr = "SORT(" & sEvalStr & ",, 1)"
    End If
    If bSorted And SortOrder = lsDesc Then
        sEvalStr = "SORT(" & sEvalStr & ",, -1)"
    End If
    
    sEvalStr = "=" & sEvalStr
    ItemsInField = this.Sheet.Evaluate(sEvalStr)
       
    'Transpose to allow referncing with single index and utilise ubound and lbound with one dimension
    On Error Resume Next
    ItemsInField = WorksheetFunction.Transpose(ItemsInField)
    If Err.Number <> 0 Then
        ItemsInField = Null
    End If
    On Error GoTo 0

End Function


Public Function FieldItemByIndex(ByVal sFieldName As String, ByVal i As Long, _
    Optional bFiltered As Boolean = False) As Variant
'Returns the i'th item (1-indexed) from field sFieldName
    
    Dim lo As ListObject

    Set lo = SelectedListObject(bFiltered)
    
    Select Case True
    Case WorksheetFunction.CountIfs(lo.HeaderRowRange, sFieldName) = 0
        FieldItemByIndex = "<<Field name out of range>>"
    Case i > Me.NumberOfRecords(bFiltered)
        FieldItemByIndex = "<<Index out of range>>"
    Case Else
        FieldItemByIndex = lo.ListColumns(sFieldName).DataBodyRange.Cells(i)
    End Select

End Function


Public Function InsertFromDictionary(InsertValues As Dictionary) As Boolean
'Insets one record from dictionary where keys are fieldheaders

    Dim key As Variant
    Dim bKeysAreValid As Boolean
    Dim i As Double

    bKeysAreValid = True

    'First test that keys in dictionary correspond to field headers
    For Each key In InsertValues.Keys
        bKeysAreValid = bKeysAreValid And _
            (WorksheetFunction.CountIfs(this.ListObj.HeaderRowRange, key) = 1)
    Next key
    
    'Insert values
    If bKeysAreValid Then
        InsertFromDictionary = True
        AddOneRowToListObject this.ListObj
        i = this.ListObj.DataBodyRange.Rows.Count
        For Each key In InsertValues.Keys
            this.ListObj.ListColumns(key).DataBodyRange(i) = InsertValues.item(key)
        Next key
    Else
        InsertFromDictionary = False
    End If


End Function


Public Function Filter(ByVal sFilterString As String, Optional bSorted As Boolean = False, _
    Optional ByVal sSortField As String, Optional ByVal SortOrder As lsSortOrder = lsAsc)
'Filters the original listobject storage into a 2nd list object
'Format of sFilterString for example paramater set "[a] = 2" translates into excel function as follows:
'=FILTER(tbl_Name, tbl_Name[a] = 2)
'For sorting above by column [b] set bSorted = True and sSortField = "[b]"

    Dim v As Variant
    Dim iSortColIndex As Integer
    
    'If Original list is blank then clear filter list and exit sub
    '(there is notging to filter)
    If this.ListObj.DataBodyRange Is Nothing Then
        If Not (this.FilteredListObj.DataBodyRange Is Nothing) Then
            this.FilteredListObj.DataBodyRange.Rows.Delete
        End If
        Exit Function
    End If
          
    On Error Resume Next
    this.FilteredListObj.DataBodyRange.Delete
    On Error GoTo 0
    
   
    sFilterString = Replace(sFilterString, "[", this.ListObj.Name & "[")
    sFilterString = "FILTER(" & this.ListObj.Name & ", " & sFilterString & ")"
    
    If bSorted Then
        sSortField = Replace(sSortField, "[", "")
        sSortField = Replace(sSortField, "]", "")
        iSortColIndex = WorksheetFunction.Match(sSortField, this.ListObj.HeaderRowRange, 0)
        sFilterString = "SORT(" & sFilterString & ", " & iSortColIndex & ", " & SortOrder & ")"
    End If
    
    
    sFilterString = "=" & sFilterString
    
    v = this.Sheet.Evaluate(sFilterString)
    
    Select Case NumberOfArrayDimensions(v)
        Case 2
            AddOneRowToListObject this.FilteredListObj
            this.FilteredListObj.DataBodyRange.Cells(1).Resize(UBound(v, 1), UBound(v, 2)) = v
        Case 1
            'If the filter returns only one row the variant is a one dimensional array and
            'requires a different mechanism to write to sheet
            AddOneRowToListObject this.FilteredListObj
            this.FilteredListObj.DataBodyRange.Cells(1).Resize(, UBound(v)) = v
        Case 0
          'No data returned.  Keep filtered list empty
    End Select

End Function


Public Sub ReplaceDataWithFilteredData()
'Replaces the Original data per the listobject with the filtered versions per This.FilteredListObject
'Filtered List object is cleared
    
    If Not this.ListObj.DataBodyRange Is Nothing Then
        this.ListObj.DataBodyRange.Rows.Delete
    End If
    
    If Not (this.FilteredListObj.DataBodyRange Is Nothing) Then
        AddOneRowToListObject this.ListObj
        this.FilteredListObj.DataBodyRange.Copy
        this.ListObj.DataBodyRange.Cells(1).PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        this.FilteredListObj.DataBodyRange.Rows.Delete
    End If
    

End Sub

Public Function Xlookup(ByVal LookupValue As Variant, _
    ByVal sLookupArray As String, _
    ByVal sReturnArray As String, _
    Optional ByVal IfNotFound As Variant, _
    Optional ByVal MatchMode As Integer = 0, _
    Optional SearchMode As Integer = 1, _
    Optional ByVal bFiltered As Boolean = False) As Variant

'Replicates Xlookup on the liststorage, either original or filtered data depending on value of
'the bFiltered parameter.  Table name does not need to be specified for example
'Xlookup([a], [b], [c]) becomes Xlookup(tbl_TableName[a], tbl_TableName[b], tbl_TableName[c])
'Null is returned when optional parameter IfNotFound is missing and xlookup is not found


    Dim sEvalStr As String
    Dim lo As ListObject
    Dim vReturnValue As Variant

    Set lo = SelectedListObject(bFiltered)

    sEvalStr = "=XLOOKUP(lookup_value, lookup_array, return_array, if_not_found," & _
        "match_mode, search_mode) "
    
    'Add quotes if values are strings
    If VarType(LookupValue) = vbString Then
        LookupValue = """" & LookupValue & """"
    End If
    If VarType(IfNotFound) = vbString Then
        IfNotFound = """" & IfNotFound & """"
    End If
    

    sEvalStr = Replace(sEvalStr, "lookup_value", LookupValue)
    sEvalStr = Replace(sEvalStr, "lookup_array", sLookupArray)
    sEvalStr = Replace(sEvalStr, "return_array", sReturnArray)
    If IsMissing(IfNotFound) Then
        sEvalStr = Replace(sEvalStr, "if_not_found", "")
    Else
        sEvalStr = Replace(sEvalStr, "if_not_found", IfNotFound)
    End If
    sEvalStr = Replace(sEvalStr, "match_mode", MatchMode)
    sEvalStr = Replace(sEvalStr, "search_mode", SearchMode)

    sEvalStr = Replace(sEvalStr, "[", lo.Name & "[")
        
    vReturnValue = this.Sheet.Evaluate(sEvalStr)
    
    'Return null if no value found and not alternative specified
    If WorksheetFunction.IsError(vReturnValue) Then
        Xlookup = Null
    Else
        Xlookup = vReturnValue
    End If

    
End Function



Function StorageAlreadyExists(ByRef wkb As Workbook, ByVal sName As String) As Boolean

    Dim sht As Worksheet
    Dim sTestSheet As String
    Dim sTestListObject As String

    'Check for existing sheet
    On Error Resume Next
    sTestSheet = wkb.Sheets(sName).Name
    If Err.Number = 0 Then
        StorageAlreadyExists = True
    End If
    
    'Check for existing listobject
    For Each sht In wkb.Worksheets
        On Error Resume Next
        sTestListObject = sht.ListObjects("tbl_" & sName)
        If Err.Number = 0 Then
            StorageAlreadyExists = True
        End If
        On Error GoTo 0
    Next sht
        

End Function


Private Function CreateListStorageSheet(ByRef wkb As Workbook, ByVal sName As String) As Worksheet

    Dim sht As Worksheet
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    sht.Name = sName
    Set CreateListStorageSheet = sht
    
End Function


Private Sub FormatListStorageSheet(ByRef sht As Worksheet)
'Applies my preferred sheet formattting

    sht.Activate
    
    sht.Cells.Font.Name = "Calibri"
    sht.Cells.Font.Size = 11
      
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 80
    sht.DisplayPageBreaks = False
    sht.Columns("A:A").ColumnWidth = 4
    
    sht.Names.Add Name:="SheetHeading", RefersTo:="=$B$2"
    sht.Names.Add Name:="SheetCategory", RefersTo:="=$A$1"
    
    With sht.Range("SheetHeading")
        .Value = sht.Name
        .Font.Bold = True
        .Font.Size = 16
    End With
    
    With sht.Range("SheetCategory")
        .Value = "List Storage"
        .Font.Color = RGB(170, 170, 170)
        .Font.Size = 8
    End With
    
    'Freeze Panes
    sht.Activate
    ActiveWindow.SplitRow = 7
    ActiveWindow.FreezePanes = True


End Sub


Private Function CreateListStorageListObject(ByRef sht As Worksheet, Headings() As String) As ListObject

    Dim i As Integer
    Const iTableRowStart As Integer = 7
    Const iTableColStart As Integer = 2
    Dim lo As ListObject
    
    For i = LBound(Headings) To UBound(Headings)
        sht.Cells(iTableRowStart, iTableColStart + i).Value = Headings(i)
    Next i

    Set lo = sht.ListObjects.Add(xlSrcRange, _
        sht.Cells(iTableRowStart, iTableColStart).CurrentRegion, , xlYes)
    lo.Name = "tbl_" & sht.Name
    
    'Set number format to text to ensure no errors when saving formulas for
    'example which will be correctly saved as strings with below setting.
    lo.Range.EntireColumn.NumberFormat = "@"
    
    sht.Cells(iTableRowStart - 2, iTableColStart).Value = "Original List"
    sht.Cells(iTableRowStart - 2, iTableColStart).Font.Bold = True
    
    Set CreateListStorageListObject = lo

End Function


Private Function CreateFilteredListStorageListObject(ByRef sht As Worksheet, Headings() As String) As ListObject

    Dim i As Integer
    Const iTableRowStart As Integer = 7
    Dim iTableColStart As Integer
    Const iUnfilteredTableColStart As Integer = 2
    Const iGapBetweenTables = 1
    Dim lo As ListObject
    
    iTableColStart = iUnfilteredTableColStart + (UBound(Headings)) + (iGapBetweenTables + 1)
    
    For i = LBound(Headings) To UBound(Headings)
        sht.Cells(iTableRowStart, iTableColStart + i).Value = Headings(i)
    Next i

    Set lo = sht.ListObjects.Add(xlSrcRange, _
        sht.Cells(iTableRowStart, iTableColStart).CurrentRegion, , xlYes)
    lo.Name = "tbl_" & sht.Name & "_Filtered"
    
    'Set number format to text to ensure no errors when saving formulas for
    'example which will be correctly saved as strings with below setting.
    'Safer to set EntireCol as values are pasted before list is expanded
    lo.Range.EntireColumn.NumberFormat = "@"
    
    sht.Cells(iTableRowStart - 2, iTableColStart).Value = "Filtered List"
    sht.Cells(iTableRowStart - 2, iTableColStart).Font.Bold = True
    
    Set CreateFilteredListStorageListObject = lo

End Function



Private Sub FormatStorageListObject(ByRef lo As ListObject)

    Dim sty As TableStyle
    Dim wkb As Workbook
    
    Set wkb = lo.Parent.Parent
    
    On Error Resume Next
    wkb.TableStyles.Add ("SpreadsheetBiStyle")
    On Error GoTo 0
    Set sty = wkb.TableStyles("SpreadsheetBiStyle")
    
    'Set Header Format
    With sty.TableStyleElements(xlHeaderRow)
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Borders.item(xlEdgeTop).LineStyle = xlSolid
        .Borders.item(xlEdgeTop).Weight = xlMedium
        .Borders.item(xlEdgeBottom).LineStyle = xlSolid
        .Borders.item(xlEdgeBottom).Weight = xlMedium
    End With

    'Set row stripe format
    sty.TableStyleElements(xlRowStripe1).Interior.Color = RGB(217, 217, 217)
    sty.TableStyleElements(xlRowStripe2).Interior.Color = RGB(255, 255, 255)
    
    'Set whole table bottom edge format
    sty.TableStyleElements(xlWholeTable).Borders.item(xlEdgeBottom).LineStyle = xlSolid
    sty.TableStyleElements(xlWholeTable).Borders.item(xlEdgeBottom).Weight = xlMedium

    
    'Apply custom style and set other attributes
    lo.TableStyle = "SpreadsheetBiStyle"
    With lo.HeaderRowRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
    End With
    
    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.EntireColumn.AutoFit
    End If
    

End Sub


Public Sub AddFieldValidationList(ByVal sFieldName As String, sValidationList As String)

    Dim TestCount As Long

    'Ensure databody range exists
    On Error Resume Next
    TestCount = this.ListObj.DataBodyRange.Rows.Count
    If Err.Number <> 0 Then AddOneRowToListObject this.ListObj

    With this.ListObj.ListColumns(sFieldName).DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=sValidationList
    End With
End Sub



Private Sub AddOneRowToListObject(lo As ListObject)

    Dim str As String
    
    On Error Resume Next
    str = lo.DataBodyRange.Address
    If Err.Number <> 0 Then
        'Force empty row in databody range if it does not yet exist
        lo.HeaderRowRange.Cells(1).Offset(1, 0) = " "
        lo.HeaderRowRange.Cells(1).Offset(1, 0).ClearContents
    Else
        lo.Resize lo.Range.Resize(lo.Range.Rows.Count + 1)
    End If
    On Error GoTo 0
    

End Sub


Private Function SelectedListObject(ByVal bFiltered) As ListObject
'Returns either This.ListObj or This.FilteredListObj depending on the
'bFiltered paramater

    If bFiltered Then
        Set SelectedListObject = this.FilteredListObj
    Else
        Set SelectedListObject = this.ListObj
    End If

End Function
 


