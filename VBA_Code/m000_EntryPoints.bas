Attribute VB_Name = "m000_EntryPoints"
Option Explicit

Sub InsertFormattedSheetIntoActiveWorkbook()
    
    Dim sht As Worksheet
    
    Set sht = ActiveWorkbook.Sheets.Add(After:=ActiveSheet)
    FormatSheet sht

End Sub





Sub AddDummyTableToDataModel()

' Rather perform below via import of power query m text file

'    ActiveWorkbook.Queries.Add _
'        Name:="DummyTable", _
'        Formula:= _
'        "#table(" & Chr(10) & "        type table[" & Chr(10) & "            #""DummyFieldName""=text " & Chr(10) & "            ], " & Chr(10) & "        {" & Chr(10) & "            {null}" & Chr(10) & "        }" & Chr(10) & "    )"
'
'    Workbooks("MasterBI.xlsm").Connections.Add2 "Query - DummyTable", _
'        "Connection to the 'DummyTable' query in the workbook.", _
'        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=DummyTable;Extended Properties=" _
'        , """DummyTable""", 6, True, False
        
End Sub

