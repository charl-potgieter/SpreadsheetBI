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
'

'Use "This" declaration as an easy way to get intellisense to the classes private variables
'https://rubberduckvba.wordpress.com/2020/02/27/vba-classes-gateway-to-solid/\
'
Private Type TypeListStorage
    wkb As Workbook
    ReportingSheet As ReportingSheet
    ListObj As ListObject
    FilteredListObj As ListObject
    CreatedFromPowerQuery As Boolean
End Type
Private this As TypeListStorage

Public Enum lsSortOrder  'ls = ListStorage
    lsAsc = 1
    lsDesc = -1
End Enum


Public Function CreateStorage(ByRef wkb As Workbook, ByVal sName As String, Headings As Variant) As Boolean
'Does nothing and returns FALSE if sheet or listobject already exist
'Creates a new sheet named "sName" and listobject named "tbl_sName" in wkb and returns TRUE
    
    If StorageAlreadyExists(wkb, sName) Then
        CreateStorage = False
    Else
        CreateStorage = True
        
        Set this.wkb = wkb
        Set this.ReportingSheet = CreateListStorageSheet(wkb, sName)
        
        Set this.ListObj = CreateListStorageListObject _
            (this.ReportingSheet.Sheet, Headings)
        FormatStorageListObject this.ListObj
        
        Set this.FilteredListObj = CreateFilteredListStorageListObject _
            (this.ReportingSheet.Sheet, Headings)
        FormatStorageListObject this.FilteredListObj
        
        
        this.FilteredListObj.Range.Columns.Group
        this.ReportingSheet.Sheet.Outline.ShowLevels ColumnLevels:=1
        
        this.CreatedFromPowerQuery = False
        
    End If

End Function


Public Function CreateStorageFromPowerQuery(ByVal wkb As Workbook, _
    ByVal sName As String, ByVal PowerQueryName As String _
    ) As Boolean

    Dim Headings As Variant

    If StorageAlreadyExists(wkb, sName) Or Not QueryExists(PowerQueryName) Then
        CreateStorageFromPowerQuery = False
    Else

        this.CreatedFromPowerQuery = True
        Set this.wkb = wkb
        Set this.ReportingSheet = CreateListStorageSheet(wkb, sName)
         
        'Output the Power Query to a worksheet table
        Set this.ListObj = this.ReportingSheet.Sheet.ListObjects.Add( _
            SourceType:=0, _
            Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & PowerQueryName & ";Extended Properties=""""", _
            Destination:=Range("$J15"))
            
        With this.ListObj
            .QueryTable.CommandType = xlCmdSql
            .QueryTable.CommandText = Array("SELECT * FROM [" & PowerQueryName & "]")
            .QueryTable.Refresh BackgroundQuery:=False
            .Name = "tbl_" & sName
        End With

        FormatStorageListObject this.ListObj

        'Double transpose trick converts to a 1 dimensional array
        'https://stackoverflow.com/questions/7649046/one-dimensional-array-from-excel-range/7651439
        Headings = WorksheetFunction.Transpose( _
            WorksheetFunction.Transpose(this.ListObj.HeaderRowRange))
            
        Set this.FilteredListObj = CreateFilteredListStorageListObject _
            (this.ReportingSheet.Sheet, Headings)
        FormatStorageListObject this.FilteredListObj
        this.FilteredListObj.Range.Columns.Group
        this.ReportingSheet.Sheet.Outline.ShowLevels ColumnLevels:=1

    End If

End Function


Public Function AssignStorage(ByRef wkb As Workbook, ByVal sName As String) As Boolean
'Assigns the storage object to a specific sheet and listobject
    
    
    On Error Resume Next
    
    Set this.wkb = wkb
    Set this.ReportingSheet = New ReportingSheet
    this.ReportingSheet.AssignExistingSheet wkb.Worksheets(sName)
    Set this.ListObj = this.ReportingSheet.Sheet.ListObjects("tbl_" & sName)
    Set this.FilteredListObj = this.ReportingSheet.Sheet.ListObjects("tbl_" & sName & "_Filtered")
    
    this.ListObj.ShowAutoFilter = True
    this.ListObj.AutoFilter.ShowAllData
    this.CreatedFromPowerQuery = (this.ListObj.SourceType = xlSrcQuery)
    
    AssignStorage = (Err.Number = 0)
    
    On Error GoTo 0

End Function


Public Property Get Name() As String
    Name = this.ReportingSheet.Name
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

Public Property Get Sheet() As Worksheet
    Set Sheet = this.ReportingSheet.Sheet
End Property


Public Sub AddBlankRow()

    If Not this.CreatedFromPowerQuery Then
        AddOneRowToListObject this.ListObj
    End If

End Sub


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
    ItemsInField = this.ReportingSheet.Sheet.Evaluate(sEvalStr)
       
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

    If this.CreatedFromPowerQuery Then
        InsertFromDictionary = False
        Exit Function
    End If

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
            this.ListObj.ListColumns(key).DataBodyRange(i) = InsertValues.Item(key)
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
    
    v = this.ReportingSheet.Sheet.Evaluate(sFilterString)
    
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


Function ReplaceDataWithFilteredData() As Boolean
'Replaces the Original data per the listobject with the filtered versions per This.FilteredListObject
'Filtered List object is cleared
    
    If this.CreatedFromPowerQuery Then
        ReplaceDataWithFilteredData = False
        Exit Function
    End If
    
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
    
    ReplaceDataWithFilteredData = True

End Function

Public Function Xlookup(ByVal LookupValue As Variant, _
    ByVal sLookupArray As String, _
    ByVal sReturnArray As String, _
    Optional ByVal IfNotFound As Variant, _
    Optional ByVal MatchMode As Integer = 0, _
    Optional SearchMode As Integer = 1, _
    Optional ByVal bFiltered As Boolean = False) As Variant

'Replicates Xlookup on the liststorage, either original or filtered data depending on value of
'the bFiltered parameter.  Table name does not need to be specified for example
'Xlookup(LookupVal, [b], [c]) becomes Xlookup(LookupVal, tbl_TableName[b], tbl_TableName[c])
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
        
    vReturnValue = this.ReportingSheet.Sheet.Evaluate(sEvalStr)
    
    'Return null if no value found and not alternative specified
    If WorksheetFunction.IsError(vReturnValue) Then
        Xlookup = Null
    Else
        Xlookup = vReturnValue
    End If

    
End Function


Public Function Delete() As Boolean

    On Error Resume Next
    this.ReportingSheet.Delete
    Delete = (Err.Number = 0)
    On Error Resume Next

End Function


Public Function ListObj() As ListObject
    Set ListObj = this.ListObj
End Function


Public Function FilteredListObj() As ListObject
    Set FilteredListObj = this.FilteredListObj
End Function



Public Function StorageAlreadyExists(ByRef wkb As Workbook, ByVal sName As String) As Boolean

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


Public Function IsEmpty(Optional ByVal bFiltered As Boolean = False) As Boolean

    Dim test As Double
    On Error Resume Next
    
    test = SelectedListObject(bFiltered).DataBodyRange.Rows.Count
    IsEmpty = (Err.Number <> 0)
    On Error GoTo 0
    

End Function


Public Function ClearData() As Boolean

    If this.CreatedFromPowerQuery Then
        ClearData = False
    End If

    On Error Resume Next 'In event there is no databody range
    this.ListObj.DataBodyRange.Rows.Delete
    On Error GoTo 0
    ClearData = True


End Function



Private Function CreateListStorageSheet(ByRef wkb As Workbook, ByVal sName As String) As ReportingSheet

    Dim RptSheet As ReportingSheet
    
    Set RptSheet = New ReportingSheet
    RptSheet.Create wkb
    RptSheet.Name = sName
    RptSheet.Category = "List Storage"
    RptSheet.Heading = "Report Properties"
    InsertIndexPage wkb
    RptSheet.Sheet.Activate
    Set CreateListStorageSheet = RptSheet
    Set RptSheet = Nothing
    
End Function




Private Function CreateListStorageListObject(ByRef sht As Worksheet, Headings As Variant) As ListObject

    Dim i As Integer
    Const iTableRowStart As Integer = 15
    Const iTableColStart As Integer = 10
    Dim lo As ListObject
    
    For i = LBound(Headings) To UBound(Headings)
        sht.Cells(iTableRowStart, iTableColStart + i - LBound(Headings)).Value = Headings(i)
    Next i

    Set lo = sht.ListObjects.Add(xlSrcRange, _
        sht.Cells(iTableRowStart, iTableColStart).CurrentRegion, , xlYes)
    lo.Name = "tbl_" & sht.Name
    
    'Set number format to text to ensure no errors when saving formulas for
    'example which will be correctly saved as strings with below setting.
    lo.Range.EntireColumn.NumberFormat = "@"
    
    sht.Cells(iTableRowStart - 2, iTableColStart).Font.Bold = True
    
    Set CreateListStorageListObject = lo

End Function


Private Function CreateFilteredListStorageListObject(ByRef sht As Worksheet, Headings As Variant) As ListObject

    Dim i As Integer
    Const iTableRowStart As Integer = 15
    Dim iTableColStart As Integer
    Const iUnfilteredTableColStart As Integer = 10
    Const iGapBetweenTables = 1
    Dim lo As ListObject
    
    iTableColStart = iUnfilteredTableColStart + (UBound(Headings)) + (iGapBetweenTables + 1)
    
    For i = LBound(Headings) To UBound(Headings)
        sht.Cells(iTableRowStart, iTableColStart + i - LBound(Headings)).Value = Headings(i)
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
        .Borders.Item(xlEdgeTop).LineStyle = xlSolid
        .Borders.Item(xlEdgeTop).Weight = xlMedium
        .Borders.Item(xlEdgeBottom).LineStyle = xlSolid
        .Borders.Item(xlEdgeBottom).Weight = xlMedium
    End With

    'Set row stripe format
    sty.TableStyleElements(xlRowStripe1).Interior.Color = RGB(217, 217, 217)
    sty.TableStyleElements(xlRowStripe2).Interior.Color = RGB(255, 255, 255)
    
    'Set whole table bottom edge format
    sty.TableStyleElements(xlWholeTable).Borders.Item(xlEdgeBottom).LineStyle = xlSolid
    sty.TableStyleElements(xlWholeTable).Borders.Item(xlEdgeBottom).Weight = xlMedium

    
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


Public Function AddFieldValidationList(ByVal sFieldName As String, sValidationList As String) As Boolean

    Dim TestCount As Long

    If this.CreatedFromPowerQuery Then
        AddFieldValidationList = False
    Else
        AddFieldValidationList = True
    End If
    

    'Ensure databody range exists
    On Error Resume Next
    TestCount = this.ListObj.DataBodyRange.Rows.Count
    If Err.Number <> 0 Then AddOneRowToListObject this.ListObj

    With this.ListObj.ListColumns(sFieldName).DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=sValidationList
    End With
End Function



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
 


Private Sub Class_Terminate()
   Set this.FilteredListObj = Nothing
   Set this.ListObj = Nothing
   Set this.ReportingSheet = Nothing
   Set this.wkb = Nothing
End Sub







