Attribute VB_Name = "Module1"
Option Explicit


Sub Test()

    
    Dim i
    
    For Each i In ActiveWorkbook.Connections
        Debug.Print (i & " " & i.Type)
    Next i
    
    
    
    


End Sub





Sub Macro2()
'
' Macro2 Macro
'

'
    Selection.ListObject.TableObject.Refresh
    Selection.ListObject.TableObject.Refresh
    Workbooks("MasterBI.xlsm").Connections.Add2 "Query - Myquery", _
        "Connection to the 'Myquery' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Myquery;Extended Properties=" _
        , """Myquery""", 6, True, False
End Sub

