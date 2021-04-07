Attribute VB_Name = "ZZZ_Testing"
Option Explicit


Sub test()

    Dim a As ListStorage
    Dim Headings(3) As String
    Dim bCreated As Boolean
    Dim d As Dictionary
    
    
    
    Set a = New ListStorage
    Set d = New Dictionary
    
    d.Add "help", "a"
    d.Add "test", "b"
    
    Headings(0) = "a"
    Headings(1) = "b"
    Headings(2) = "c"
    Headings(3) = "d"
    
    bCreated = a.CreateStorage(ActiveWorkbook, "Test2", Headings)
    
    
End Sub

Sub Test2()

    Dim ls As ListStorage
    Dim b As Boolean
    Dim d As Dictionary
    
    Set ls = New ListStorage
    Set d = New Dictionary
    
    d.Add "c", 4
    d.Add "a", 2
    d.Add "b", 7
    d.Add "d", "Excellent!"
     
    
    b = ls.AssignStorage(ActiveWorkbook, "Test")
    b = ls.InsertFromDictionary(d)
    
    
    Debug.Print (b)
    


End Sub





Sub test4()

    Dim a(2) As String
    Dim b(2) As String
    Dim dic As Dictionary

    a(0) = 1
    a(1) = 2
    a(2) = 3
    b(0) = 10
    b(1) = 11
    b(2) = 12

    Set dic = New Dictionary
    dic.Add "a", a
    dic.Add "b", b



End Sub


Sub test5()

    Dim lo As ListObject
    
    Set lo = ActiveSheet.ListObjects(1)
    lo.ShowAutoFilter = True
    lo.AutoFilter.ShowAllData


End Sub


Sub Test6()

    Dim v() As Variant
    
    v = Evaluate("=UNIQUE(FILTER(tbl_Test[b], tbl_Test[b]<>""""))")
    ReDim Preserve v(2, 2)

End Sub

Sub Test7()

    Dim ls As ListStorage
    Dim v As Variant
    Dim i As Long
    
    Set ls = New ListStorage

    
    ls.AssignStorage ActiveWorkbook, "Test2"

    v = ls.ItemsInField("a", bIgnoreBlanks:=False, bUnique:=True, bSorted:=True, SortOrder:=lsAsc, bFiltered:=False)
    
    ActiveSheet.Range("Q8").Resize(UBound(v)) = WorksheetFunction.Transpose(v)

End Sub


Sub Test8()

    Dim v As Variant
    Dim v2 As Variant
    
    v = Evaluate("=FILTER(tbl_Test, tbl_Test[a]=22)")
    v2 = Evaluate("=FILTER(tbl_Test, tbl_Test[a]=222)")
    ActiveSheet.ListObjects("Table1").ListColumns(1).DataBodyRange.Cells(1).Resize(UBound(v2, 1), UBound(v2, 2)).Value = v2
    

End Sub


Sub Test9()

    Dim lo As ListObject
    
    Set lo = ActiveSheet.ListObjects("tbl_Test_Filtered")
    
    

End Sub



Sub Test10()

    Dim a As ListStorage
    Dim Headings(3) As String
    Dim bCreated As Boolean
    Dim d As Dictionary
    
    
    
    Set a = New ListStorage
    Set d = New Dictionary
    
    d.Add "help", "a"
    d.Add "test", "b"
    
    Headings(0) = "a"
    Headings(1) = "b"
    Headings(2) = "c"
    Headings(3) = "d"
    
    bCreated = a.CreateStorage(ActiveWorkbook, "Test2", Headings)
    
    
End Sub


Sub Test11()

    Dim a As ListStorage
    
    
    Set a = New ListStorage
    
    a.AssignStorage ActiveWorkbook, "Test2"
    
    a.Filter "([a] = ""ace"")"
    
    Debug.Print (a.FieldItemByIndex("b", 4))
    Debug.Print (a.FieldItemByIndex("b", 4, False))
    Debug.Print (a.FieldItemByIndex("b", 4, True))
    
End Sub




Sub Test12()

    Dim a As ListStorage
    
    Set a = New ListStorage
    
    a.AssignStorage ActiveWorkbook, "Test2"
    
    Debug.Print (a.Xlookup(3, "[a]", "[c]", "blah", , , True))
        
        
    
End Sub

Sub Test13()

    Dim a As ListStorage
    
    Set a = New ListStorage
    
    a.AssignStorage ActiveWorkbook, "Test2"
        
    a.ReplaceDataWithFilteredData
    
End Sub




Sub TestEmptyStorageFilter()

    Dim ls As ListStorage
    Dim sHeadings(2) As String
    Dim bStorageCreated As Boolean
    
    sHeadings(0) = "a"
    sHeadings(1) = "b"
    sHeadings(2) = "c"
    
    Set ls = New ListStorage
    bStorageCreated = ls.CreateStorage(ThisWorkbook, "Test4", sHeadings)
    
    If Not bStorageCreated Then
        ls.AssignStorage ActiveWorkbook, "Test4"
    End If
    
    
    ls.Filter ("[a] = 77")
    

End Sub



Sub TestArray()
    
    Dim v() As Variant
    Dim item As Variant
    
    v = Array(1, 2, 3)
    
    For Each item In v
        Debug.Print item
    Next item



End Sub

'Sub TestCateogories()
'
'    Dim v As Variant
'    Dim item As Variant
'
'    v = m001_DataPivotReporting.ReadUniquePivotReportCategories
'    For Each item In v
'        Debug.Print (item)
'    Next item
'
'
'End Sub




'Sub TestUserFrm()
'
'    Dim v As Variant
'    Dim item As Variant
'    Dim uf As ufPivotReportGenerator
'
'
'    Set uf = New ufPivotReportGenerator
'    v = m001_DataPivotReporting.ReadUniquePivotReportCategories
'    For Each item In v
'        uf.lbCategories.AddItem item
'    Next item
'
'    uf.Show
'
'    Set uf = Nothing
'
'End Sub



Sub TestFreeze()

    
    If (ActiveWindow.SplitColumn <> 0 And CStr(ActiveWindow.SplitColumn) <> "") Or _
        (ActiveWindow.SplitRow <> 0 And CStr(ActiveWindow.SplitRow) <> "") Then
        MsgBox ("Split")
    Else
        MsgBox ("Not")
    End If
        

End Sub


Sub TestReadFreezePanes()

    Dim sSheetName As String
    Dim ls As ListStorage
    Set ls = New ListStorage
    Dim sFreezePosition As String
    Dim vReturnValue As Variant
    
    sSheetName = "Pvt_b"
    ls.AssignStorage ActiveWorkbook, "ReportSheetProperties"
    
    vReturnValue = ls.Xlookup(sSheetName & "ViewLayoutDataType" & "FreezePanes", _
        "[SheetName] & [DataType] & [Property]", _
        "[Value]")

    
    If IsNull(vReturnValue) Then
        MsgBox ("Null")
    Else
        MsgBox (vReturnValue)
    End If
    
    
End Sub


Sub PivotCopyFormatValues()
'select pivot table cell first

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim rngPT As Range
    Dim rngPTa As Range
    Dim rngCopy As Range
    Dim rngCopy2 As Range
    Dim lRowTop As Long
    Dim lRowsPT As Long
    Dim lRowPage As Long
    Dim msgSpace As String
    
    On Error Resume Next
    Set pt = ActiveCell.PivotTable
    Set rngPTa = pt.PageRange
    On Error GoTo errHandler
    
    If pt Is Nothing Then
        MsgBox "Could not copy pivot table for active cell"
        GoTo exitHandler
    End If
    
    If pt.PageFieldOrder = xlOverThenDown Then
      If pt.PageFields.Count > 1 Then
        msgSpace = "Horizontal filters with spaces." _
          & vbCrLf _
          & "Could not copy Filters formatting."
      End If
    End If
    
    Set rngPT = pt.TableRange1
    lRowTop = rngPT.Rows(1).Row
    lRowsPT = rngPT.Rows.Count
    Set ws = Worksheets.Add
    Set rngCopy = rngPT.Resize(lRowsPT - 1)
    Set rngCopy2 = rngPT.Rows(lRowsPT)
    
    rngCopy.Copy Destination:=ws.Cells(lRowTop, 1)
    rngCopy2.Copy _
      Destination:=ws.Cells(lRowTop + lRowsPT - 1, 1)
    
    If Not rngPTa Is Nothing Then
        lRowPage = rngPTa.Rows(1).Row
        rngPTa.Copy Destination:=ws.Cells(lRowPage, 1)
    End If
        
    ws.Columns.AutoFit
    If msgSpace <> "" Then
      MsgBox msgSpace
    End If
    
exitHandler:
        Exit Sub
errHandler:
        MsgBox "Could not copy pivot table for active cell"
        Resume exitHandler
End Sub

'Sub TestPivotReportValueCopy()
'
'    Dim pr As PivotReport
'    Dim sTest As String
'    Dim bAssignedOk As Boolean
'    Dim wkb As Workbook
'    Dim sht As Worksheet
'
'    Set pr = New PivotReport
'    bAssignedOk = pr.AssignToExistingSheet(ActiveSheet)
'
'    If Not bAssignedOk Then
'        MsgBox ("Not a valid Power Pivot Report sheet")
'        Exit Sub
'    End If
'
'    Set wkb = Workbooks.Add
'    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
'    sht.Name = "Blah"
'    sht.Cells.Font.Name = "Calibri"
'    sht.Cells.Font.Size = 11
'
'    'Add heading and category
'    sht.Names.Add Name:="SheetCategory", RefersTo:="=$A$1"
'    sht.Names.Add Name:="SheetHeading", RefersTo:="=$B$2"
'    sht.Range("SheetHeading") = "Blah"
'
'    sht.Range("SheetCategory") = "Blah"
'        With sht.Range("SheetHeading")
'        .Font.Bold = True
'        .Font.Size = 16
'    End With
'
'    With sht.Range("SheetCategory")
'        .Font.Color = RGB(170, 170, 170)
'        .Font.Size = 8
'    End With
'
'    sht.Activate
'    ActiveWindow.DisplayGridlines = False
'    ActiveWindow.Zoom = 80
'    sht.DisplayPageBreaks = False
'    sht.Columns("A:A").ColumnWidth = 4
'
'
'End Sub



Sub TestPivotReportValueCopy_Alternative()

    Dim pvt As PivotTable
    
    Set pvt = ActiveSheet.PivotTables(1)
    pvt.TableRange2.ClearContents
End Sub


Sub TestForFilters()

    Dim PvtCubeField As CubeField
    
    
    For Each PvtCubeField In ActiveSheet.PivotTables(1).CubeFields
        Debug.Print PvtCubeField.Name & " : " & PvtCubeField.AllItemsVisible
    Next PvtCubeField
    

End Sub


Sub TestFilters2()

    Dim item As Variant
    
    
    For Each item In ActiveSheet.PivotTables("PivotTable1").PivotFields("[DimAccounts].[Account].[Account]").HiddenItemsList
        Debug.Print item
    Next item

End Sub




Sub TestDeleteTableConnection()

   Dim lo As ListObject
   
   Set lo = ActiveCell.ListObject

    lo.TableObject.Delete

End Sub


Sub TestActiveWorkbookChange()

    Dim ls As ListStorage
    Dim v As Variant
    Dim item As Variant
    Dim d As Dictionary
    
    Set ls = New ListStorage
    ls.AssignStorage ActiveWorkbook, "ReportSheetProperties"

    Set d = New Dictionary
    d.Add "Name", "blah blah"
    d.Add "DataType", "blah blah"

     Workbooks("Book1").Activate
    Debug.Print (ls.Xlookup("Caption", "[Property]", "[Value]"))
     
'    v = ls.ItemsInField("DataType", , True, True)
'    For Each item In v
'        Debug.Print (item)
'    Next item
    'Debug.Print (ls.FieldItemByIndex("DataType", 2))
    'ls.InsertFromDictionary d
    'ls.Filter "[Property] = ""Caption""", True, "Name", lsDesc

End Sub

'
'Sub TestStandardPivot_1()
'
'    Dim SPR As StandardPivotReport
'
'    Set SPR = New StandardPivotReport
'    SPR.DaxQueryFilePath = "C:\Users\charl\Dropbox\Dropbox_Charl\Computer_Technical\Programming_GitHub\SpreadsheetBI\TestDaxQuery_1.dax"
'    SPR.GenerateDataFromQuery ActiveWorkbook
'
'    SPR.AddCalculatedTableColumn "Double Fcst", "=[@[Forecast Amt]] * 2"
'
'End Sub






