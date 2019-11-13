Attribute VB_Name = "Module1"
Option Explicit


Sub Test()

    
    Dim i
    
    For Each i In ActiveWorkbook.Connections
        Debug.Print (i & " " & i.Type)
    Next i
    
    
    
    


End Sub


