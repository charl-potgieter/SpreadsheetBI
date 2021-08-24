Attribute VB_Name = "ZZZ_Testing"
Option Explicit


Sub TestAddQueriesToPowerPivot()

    Dim qry As WorkbookQuery
    Dim wkb As Workbook
        
    Set wkb = ActiveWorkbook
    
    For Each qry In wkb.Queries
    
        wkb.Connections.Add2 _
            Name:="Query - CalculationSource", _
            Description:="Connection to the 'CalculationSource' query in the workbook.", _
            ConnectionString:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=CalculationSource;Extended Properties=", _
            CommandText:="""" & qry.Name & """", _
            lCmdtype:=xlCmdTableCollection, _
            CreateModelConnection:=True, _
            ImportRelationShips:=False
    
    Next qry



'    Workbooks("Book3").Connections.Add2 "Query - CalculationSource", _
'        "Connection to the 'CalculationSource' query in the workbook.", _
'        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=CalculationSource;Extended Properties=" _
'        , """CalculationSource""", 6, True, False



End Sub



Sub TestLoadQueryToTabl()

    Dim cn As WorkbookConnection
    Dim qry As WorkbookQuery
    
    For Each cn In ActiveWorkbook.Connections
        If cn.Type = xlConnectionTypeOLEDB Then
            Debug.Print cn.Name & " " & cn.Type & cn.OLEDBConnection.CommandText
        End If
    Next cn
    
'
'    For Each qry In ActiveWorkbook.Queries
'        Debug.Print (qry)
'    Next qry

End Sub

