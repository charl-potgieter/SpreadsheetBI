Attribute VB_Name = "ZZZ_Testing"
Option Explicit



Sub TestHideDetails()

    Dim pvt As PivotTable
    Dim cb As CubeField
    Dim pf As PivotField
    Dim pi As PivotItem
    
    Set pvt = ActiveSheet.PivotTables(1)
    Set cb = pvt.CubeFields("[DateTable].[MonthOfYear]")
    Set pf = cb.PivotFields(1)
    Set pi = pf.PivotItems("[DateTable].[MonthOfYear].&[2]")
    pi.DrilledDown = False
    


End Sub









