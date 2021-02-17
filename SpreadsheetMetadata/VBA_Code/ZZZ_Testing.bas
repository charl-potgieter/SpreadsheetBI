Attribute VB_Name = "ZZZ_Testing"
Option Explicit


Sub Test()

    Dim a As ListStorage
    Dim Headings(3) As String
    Dim bCreated As Boolean
    
    Set a = New ListStorage
    
    Headings(0) = "a"
    Headings(1) = "b"
    Headings(2) = "c"
    Headings(3) = "d"
    
    bCreated = a.CreateStorage(ActiveWorkbook, "Test", Headings)
    
    If bCreated Then
        MsgBox (a.Name)
    Else
        MsgBox ("Not created")
    End If

    
End Sub
