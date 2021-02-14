Attribute VB_Name = "ZZZ_Testing"
Option Explicit





Sub temp()
'Requires reference to Microsoft Scripting Runtime

    Dim dict As Dictionary
    Set dict = New Dictionary
    
    dict.Add "key1", "value1"
    dict.Add "key2", "value2"
    
    Dim key As Variant
    For Each key In dict.Keys
        Debug.Print "Key: " & key, "Value: " & dict.item(key)
    Next

        

End Sub


Sub temp2()
    
    Dim a As Variant
    Dim b As Variant
    Dim v As Variant
    Dim l As Variant
    
    
    a = WorksheetFunction.Unique(Selection)
    l = ActiveSheet.ListObjects(1).DataBodyRange.Value
    
    v = ActiveSheet.ListObjects(1).ListColumns("Name").DataBodyRange.Value
    
    b = Evaluate("=FILTER(tbl_ReportProperties, tbl_ReportProperties[DataType] = ""SheetProperty"")")
    
    

End Sub


Sub temp3()

    Dim lo As ListObject
    Dim vPivotFieldArray As Variant
    
    Set lo = ActiveWorkbook.Sheets("ReportSheetProperties").ListObjects("tbl_ReportProperties")
    vPivotFieldArray = WorksheetFunction.Unique(lo.ListColumns("Name").DataBodyRange.SpecialCells(xlCellTypeVisible))
        


End Sub

Sub AddField()
    
    Dim pvt As PivotTable
    Dim pvtField As PivotField
    Set pvt = ActiveSheet.PivotTables(1)

    pvt.CubeFields("[DimAccounts].[Group]").Orientation = xlRowField
    '    .Position = 1
    
    
    'ActiveSheet.PivotTables("PivotTable6").AddDataField ActiveSheet.PivotTables( _
     '   "PivotTable6").PivotFields("B"), "Sum of B", xlSum
End Sub


    
    

