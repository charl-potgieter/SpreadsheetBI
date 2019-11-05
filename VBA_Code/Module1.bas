Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveWorkbook.Queries.Add Name:="DummyTable", Formula:= _
        "#table(" & Chr(10) & "        type table[" & Chr(10) & "            #""DummyFieldName""=text " & Chr(10) & "            ], " & Chr(10) & "        {" & Chr(10) & "            {null}" & Chr(10) & "        }" & Chr(10) & "    )"
    Workbooks("MasterBI.xlsm").Connections.Add2 "Query - DummyTable", _
        "Connection to the 'DummyTable' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=DummyTable;Extended Properties=" _
        , """DummyTable""", 6, True, False
End Sub
