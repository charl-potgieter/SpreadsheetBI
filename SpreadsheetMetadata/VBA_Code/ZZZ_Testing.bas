Attribute VB_Name = "ZZZ_Testing"
Option Explicit


Sub Test()

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


Sub Test5()

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



Sub TestArray2()

    Dim pr As PowerReport



End Sub
