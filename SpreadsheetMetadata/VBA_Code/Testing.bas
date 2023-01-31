Attribute VB_Name = "Testing"
'@Folder "Testing"
Option Explicit

Sub TestFormattingObject()

    Dim UserSelection As DisplayObject
    Set UserSelection = New DisplayObject

    

        
End Sub




Sub TestTableStyle()

    Dim CustomTableStyle As TableStyle
    
    For Each CustomTableStyle In ActiveWorkbook.TableStyles
        Debug.Print CustomTableStyle.Name & " " & CustomTableStyle.ShowAsAvailableTableStyle
    Next CustomTableStyle
    
End Sub



Sub Macro1()
    
End Sub



Sub TestApplyStyle()
    
    Dim tbl As ListObject

    Set tbl = ActiveCell.ListObject
    
    tbl.TableStyle = Workbooks("Book6").TableStyles("TempTableStyle")


End Sub





